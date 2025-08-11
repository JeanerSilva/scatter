import "./../style/visual.less";
import * as d3 from "d3";
import powerbi from "powerbi-visuals-api";

import IVisual = powerbi.extensibility.visual.IVisual;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import DataView = powerbi.DataView;
import IVisualHost = powerbi.extensibility.IVisualHost;

interface Point {
  x: number;
  y: number;
  cat?: string;
  tooltip?: string;
}

/** Quebra texto em múltiplas linhas dentro de um <text> usando <tspan>. */
function wrapText(
  textSel: d3.Selection<SVGTextElement, unknown, null, undefined>,
  maxWidthPx: number
) {
  textSel.each(function () {
    const text = d3.select<SVGTextElement, unknown>(this);
    const original = (text.text() || "").trim();
    if (!original) return;

    const words = original.split(/\s+/).reverse(); // pilha
    let line: string[] = [];
    let lineNumber = 0;
    const lineHeightEm = 1.1; // espaçamento entre linhas em "em"
    const y = text.attr("y");
    const x = text.attr("x") || "0";
    const dy0 = parseFloat(text.attr("dy") || "0");

    text.text(null); // limpa conteúdo
    let tspan = text
      .append("tspan")
      .attr("x", x)
      .attr("y", y)
      .attr("dy", dy0 + "em");

    let word: string | undefined;
    while ((word = words.pop())) {
      line.push(word);
      tspan.text(line.join(" "));
      // se estourou a largura, quebra antes da última palavra
      if (tspan.node() && tspan.node()!.getComputedTextLength() > maxWidthPx) {
        line.pop();
        tspan.text(line.join(" "));
        line = [word];
        tspan = text
          .append("tspan")
          .attr("x", x)
          .attr("y", y)
          .attr("dy", (++lineNumber * lineHeightEm + dy0) + "em")
          .text(word);
      }
    }
  });
}

export class Visual implements IVisual {
  private host?: IVisualHost;

  private svg: d3.Selection<SVGSVGElement, unknown, null, undefined>;
  private root: d3.Selection<SVGGElement, unknown, null, undefined>;
  private pointsG: d3.Selection<SVGGElement, unknown, null, undefined>;
  private lineG: d3.Selection<SVGGElement, unknown, null, undefined>;
  private xAxisG: d3.Selection<SVGGElement, unknown, null, undefined>;
  private yAxisG: d3.Selection<SVGGElement, unknown, null, undefined>;
  private xLabel: d3.Selection<SVGTextElement, unknown, null, undefined>;
  private yLabel: d3.Selection<SVGTextElement, unknown, null, undefined>;
  private eqLabel: d3.Selection<SVGTextElement, unknown, null, undefined>;

  constructor(options?: powerbi.extensibility.visual.VisualConstructorOptions) {
    this.host = options?.host;

    const container = options?.element ?? document.createElement("div");
    this.svg = d3.select(container).append("svg");
    this.root = this.svg.append("g").attr("class", "root");
    this.pointsG = this.root.append("g").attr("class", "points");
    this.lineG = this.root.append("g").attr("class", "trendline");
    this.xAxisG = this.root.append("g").attr("class", "x-axis");
    this.yAxisG = this.root.append("g").attr("class", "y-axis");
    this.xLabel = this.root.append("text").attr("class", "x-label").attr("text-anchor", "middle");
    this.yLabel = this.root.append("text").attr("class", "y-label").attr("text-anchor", "middle").attr("transform", "rotate(-90)");
    this.eqLabel = this.root.append("text").attr("class", "eq-label").attr("text-anchor", "middle");
  }

  public update(options: VisualUpdateOptions): void {
    const viewport = options.viewport;
    this.svg.attr("width", viewport.width).attr("height", viewport.height);

    const dv: DataView | undefined = options.dataViews && options.dataViews[0];
    if (!dv || !dv.categorical) {
      this.root.selectAll("*").remove();
      return;
    }

    const cat = dv.categorical.categories && dv.categorical.categories[0];
    const vals = dv.categorical.values;
    if (!cat || !vals || vals.length < 2) {
      this.root.selectAll("*").remove();
      return;
    }

    const xCol = vals[0];
    const yCol = vals[1];
    const tooltipCols = vals.slice(2);

    const catValues = cat.values;
    const xValues = xCol.values as (number | null | undefined)[];
    const yValues = yCol.values as (number | null | undefined)[];

    const n = Math.min(catValues.length, xValues.length, yValues.length);
    const points: Point[] = [];
    for (let i = 0; i < n; i++) {
      const xv = Number(xValues[i]);
      const yv = Number(yValues[i]);
      if (Number.isFinite(xv) && Number.isFinite(yv)) {
        const catTxt = String(catValues[i]);
        const pieces: string[] = [];
        pieces.push(catTxt);
        pieces.push(`${xCol.source.displayName ?? "X"}: ${xv}`);
        pieces.push(`${yCol.source.displayName ?? "Y"}: ${yv}`);
        for (const tc of tooltipCols) {
          const v = tc.values[i];
          pieces.push(`${tc.source.displayName ?? "Valor"}: ${v}`);
        }
        points.push({ x: xv, y: yv, cat: catTxt, tooltip: pieces.join("\n") });
      }
    }

    if (points.length === 0) {
      this.root.selectAll("*").remove();
      return;
    }

    // Painel de formatação
    const objects = dv.metadata.objects || {};
    const pointSize = (objects as any)?.points?.size != null ? Number((objects as any).points.size) : 50;
    const pointColor = (objects as any)?.points?.color?.solid?.color ?? "#1f77b4";
    const showTrend = (objects as any)?.trendline?.show != null ? Boolean((objects as any).trendline.show) : true;
    const lineColor = (objects as any)?.trendline?.color?.solid?.color ?? "#d62728";
    const lineWidth = (objects as any)?.trendline?.strokeWidth != null ? Number((objects as any).trendline.strokeWidth) : 1.5;
    const xTitle = (objects as any)?.axes?.xTitle ?? (xCol.source.displayName ?? "");
    const yTitle = (objects as any)?.axes?.yTitle ?? (yCol.source.displayName ?? "");

    // Margens: espaço extra para rótulos e equação
    const margin = { top: 10, right: 10, bottom: 72, left: 72 };
    const w = Math.max(0, viewport.width - margin.left - margin.right);
    const h = Math.max(0, viewport.height - margin.top - margin.bottom);
    this.root.attr("transform", `translate(${margin.left},${margin.top})`);

    // Escalas
    const x = d3.scaleLinear()
      .domain(d3.extent(points, d => d.x) as [number, number])
      .nice()
      .range([0, w]);
    const y = d3.scaleLinear()
      .domain(d3.extent(points, d => d.y) as [number, number])
      .nice()
      .range([h, 0]);

    // Eixos e rótulos (desenhar eixos primeiro)
    this.xAxisG.attr("transform", `translate(0,${h})`).call(d3.axisBottom(x));
    this.yAxisG.call(d3.axisLeft(y));

    // Rótulo X (pode usar wrap se quiser — coloque um width maior que 0)
    this.xLabel.attr("x", w / 2).attr("y", h + 36).attr("dy", "0").text(xTitle);
    // wrapText(this.xLabel, Math.min(300, w)); // habilite se quiser quebra no eixo X

    // Rótulo Y com wrap — largura máxima ~ 120 px costuma ficar bom
    this.yLabel.attr("x", -h / 2).attr("y", -56).attr("dy", "0").text(yTitle);
    wrapText(this.yLabel, 180);

    // Pontos
    const circles = this.pointsG.selectAll<SVGCircleElement, Point>("circle").data(points, (_d, i) => i as any);
    circles.exit().remove();
    const enterCircles = circles.enter().append<SVGCircleElement>("circle");
    enterCircles.attr("r", Math.sqrt(pointSize / Math.PI)).attr("fill", pointColor);
    const allCircles = enterCircles.merge(circles);
    allCircles
      .attr("cx", d => x(d.x))
      .attr("cy", d => y(d.y))
      .attr("r", Math.sqrt(pointSize / Math.PI))
      .attr("fill", pointColor);
    allCircles.each(function (d) {
      const sel = d3.select<SVGCircleElement, Point>(this);
      let titleSel = sel.select<SVGTitleElement>("title");
      if (titleSel.empty()) {
        const t = document.createElementNS("http://www.w3.org/2000/svg", "title");
        this.appendChild(t);
        titleSel = sel.select<SVGTitleElement>("title");
      }
      titleSel.text(d.tooltip ?? "");
    });

    // ===== Regressão via Pearson (Excel/Power BI): r -> R² =====
    this.lineG.selectAll("*").remove();
    this.eqLabel.text("");

    if (showTrend && points.length >= 2) {
      const xs = points.map(p => p.x);
      const ys = points.map(p => p.y);
      const nPts = xs.length;

      const meanX = d3.mean(xs) as number;
      const meanY = d3.mean(ys) as number;

      const dx = xs.map(v => v - meanX);
      const dy = ys.map(v => v - meanY);

      const Sxx = d3.sum(dx, d => d * d);
      const Syy = d3.sum(dy, d => d * d);
      const Sxy = d3.sum(dx.map((d, i) => d * dy[i]));

      const r = (Sxx > 0 && Syy > 0) ? (Sxy / Math.sqrt(Sxx * Syy)) : 0;
      const r2 = r * r;

      // Desvios padrão amostrais (Excel STDEV.S)
      const sdX = Math.sqrt(Sxx / (nPts - 1));
      const sdY = Math.sqrt(Syy / (nPts - 1));

      // Slope/intercept a partir de r (equivalente ao OLS com intercepto)
      const m = (sdX > 0) ? r * (sdY / sdX) : 0;
      const b = meanY - m * meanX;

      // Linha
      const xMin = x.domain()[0];
      const xMax = x.domain()[1];
      const y1 = m * xMin + b;
      const y2 = m * xMax + b;

      this.lineG.append("line")
        .attr("x1", x(xMin)).attr("y1", y(y1))
        .attr("x2", x(xMax)).attr("y2", y(y2))
        .attr("stroke", lineColor)
        .attr("stroke-width", lineWidth);

      // Equação + R² posicionado abaixo do rótulo X
      const mFmt = Math.abs(m) >= 100 ? m.toFixed(0) : Math.abs(m) >= 10 ? m.toFixed(2) : m.toFixed(3);
      const bFmtAbs = Math.abs(b) >= 100 ? Math.abs(b).toFixed(0) : Math.abs(b) >= 10 ? Math.abs(b).toFixed(2) : Math.abs(b).toFixed(3);
      const sign = b >= 0 ? "+" : "−";
      const eq = `y = ${mFmt}x ${sign} ${bFmtAbs}   (R² = ${r2.toFixed(3).replace(".", ",")})`;

      this.eqLabel.attr("x", w / 2).attr("y", h + 56).text(eq);

      // Logs de depuração (abra F12 no pbiviz start para ver)
      // eslint-disable-next-line no-console
      console.group("Pearson/OLS breakdown");
      // eslint-disable-next-line no-console
      console.log({ nPts, meanX, meanY, Sxx, Syy, Sxy, r, r2, sdX, sdY, slope: m, intercept: b });
      // eslint-disable-next-line no-console
      console.groupEnd();
    }
  }
}
