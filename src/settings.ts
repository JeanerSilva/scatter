import powerbi from "powerbi-visuals-api";
import DataViewObjects = powerbi.DataViewObjects;

export class VisualSettings {
  points = { size: 50, color: "#1f77b4" };
  trendline = { show: true, color: "#d62728", strokeWidth: 1.5 };
  axes = { xTitle: "", yTitle: "" };

  static getValue<T>(objects: DataViewObjects, object: string, prop: string, defaultValue: T): T {
    const o = objects && (objects as any)[object];
    const v = o && o[prop];
    if (v && typeof v === "object" && "solid" in v) {
      return (v as any).solid.color ?? defaultValue;
    }
    return (v ?? defaultValue) as T;
  }
}
