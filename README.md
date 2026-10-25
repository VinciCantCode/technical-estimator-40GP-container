# Precise Pipe — 40GP Container Pipe Loading & Commercial Estimator Model

An engineering and commercial estimation model that calculates the exact **number of pieces (`pcs`)** and **net weight in metric tons (`MT`)** of steel pipes that fit into a **40' General Purpose (40GP) container** across all standard **Nominal Pipe Sizes (NPS 1/8" to 48")** and **Pipe Schedules (5S to XXS)**.

Built for **Carbon Steel (ASME B36.10M)** and **Stainless Steel 304 / 316 (ASME B36.19M)**.

---

## Key Deliverables

1. **Linked Executive Excel Model**: [`Precise_Pipe_40GP_Container_Model_V2_Linked.xlsx`](./Precise_Pipe_40GP_Container_Model_V2_Linked.xlsx)
   * **Tab 1 (`1_Container_Calculator`)**: Interactive lookup dashboard with dropdown selectors for Material Grade, Nominal Pipe Size (NPS), Pipe Schedule, Single Pipe Length (`5.80m`, `6.00m`, `11.80m`, `12.00m`), and Cross-Section Stacking Mode (`Staggered Hexagonal` vs. `Square Grid`). Outputs 40GP Total Pieces, Metric Tons, Loading Bottleneck (`Weight Limited` vs. `Volume Limited`), unit freight economics (`$/MT` and `$/m`), and automated Pipe Nesting (telescoping) recommendations.
   * **Tab 2 (`2_Selectable_Master_Table`)**: Core calculation engine and filterable master table covering all **1,176 valid Material × NPS × Schedule combinations**.
   * **Tab 3 (`3_ASME_Spec_DB`)**: Authoritative ASME B36.10M and ASME B36.19M Outer Diameter (OD), Wall Thickness (WT), and Material Density reference database in both millimeters (`mm`) and inches (`in`).

2. **Interactive 2D Web Visualizer**: [`Precise_Pipe_40GP_Visualizer.html`](./Precise_Pipe_40GP_Visualizer.html) / [`index.html`](./index.html)
   * Zero-dependency HTML5 Canvas simulator rendering a real-time **2D Cross-Section End View** and **Lengthwise Side Elevation** of the 40GP container.
   * Visually differentiates loaded steel pipes from unloaded overweight slots, displays the dynamic payload cutoff line, and previews **Pipe Nesting (Telescoping)** inside volume-limited large-bore pipes.

---

## Engineering Methodology

* **Linear Unit Weight ($\text{kg/m}$)**:
  $$\text{Unit Weight} = (OD_{\text{mm}} - WT_{\text{mm}}) \times WT_{\text{mm}} \times C, \quad C = \pi \times 10^{-6} \times \rho_{\text{kg/m}^3}$$
  * Carbon Steel ($\rho = 7.85\text{ g/cm}^3$): $C = 0.0246615$
  * Stainless Steel 304 ($\rho = 7.93\text{ g/cm}^3$): $C = 0.0249128$
  * Stainless Steel 316 ($\rho = 7.98\text{ g/cm}^3$): $C = 0.0250699$
* **Cross-Section Geometry**:
  * Supports both **Staggered (Hexagonal Nesting)** with vertical layer pitch $\frac{\sqrt{3}}{2} \times OD \approx 0.8660254 \times OD$ and **Square Grid** stacking within configurable effective container dimensions.
* **Dual-Constraint Bottleneck Optimization**:
  $$\text{Actual Pieces} = \min\left(\text{Max Pieces by Volume}, \text{Max Pieces by Weight}\right)$$
