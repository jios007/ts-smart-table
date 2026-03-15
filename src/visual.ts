/* eslint-disable */
"use strict";

import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import DataView = powerbi.DataView;

interface TableSettings {
    headerColor: string;
    headerTextColor: string;
    rowColor1: string;
    rowColor2: string;
    textColor: string;
    fontSize: number;
    rowHeight: number;
    showCards: boolean;
    cardBackgroundColor: string;
    cardTextColor: string;
}

interface ColumnInfo {
    name: string;
    index: number;
    isNumeric: boolean;
    total: number;
}

export class TSSmartTable implements IVisual {
    private host: IVisualHost;
    private container: HTMLElement;
    private settings: TableSettings;
    private sortColumn: number = -1;
    private sortAscending: boolean = true;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.container = options.element as HTMLElement;
        this.container.style.overflow = "auto";
        this.settings = this.defaultSettings();
    }

    private defaultSettings(): TableSettings {
        return {
            headerColor: "#1E3246",
            headerTextColor: "#ffffff",
            rowColor1: "#f7fafc",
            rowColor2: "#edf2f7",
            textColor: "#2d3748",
            fontSize: 12,
            rowHeight: 32,
            showCards: true,
            cardBackgroundColor: "#1E3246",
            cardTextColor: "#ffffff"
        };
    }

    private parseSettings(dataView: DataView): void {
        if (!dataView?.metadata?.objects) return;
        const objects = dataView.metadata.objects;

        const ts = objects["tableSettings"] as any;
        const sc = objects["summaryCards"] as any;

        if (ts) {
            if (ts["headerColor"]?.solid?.color) this.settings.headerColor = ts["headerColor"].solid.color;
            if (ts["rowColor1"]?.solid?.color) this.settings.rowColor1 = ts["rowColor1"].solid.color;
            if (ts["rowColor2"]?.solid?.color) this.settings.rowColor2 = ts["rowColor2"].solid.color;
            if (ts["fontSize"] !== undefined) this.settings.fontSize = ts["fontSize"] as number;
        }
        if (sc) {
            if (sc["showCards"] !== undefined) this.settings.showCards = sc["showCards"] as boolean;
        }
    }

    public update(options: VisualUpdateOptions): void {
        const dataView = options.dataViews?.[0];
        const width = options.viewport.width;
        const height = options.viewport.height;

        this.container.innerHTML = "";
        this.container.style.width = width + "px";
        this.container.style.height = height + "px";
        this.container.style.backgroundColor = this.settings.rowColor1;
        this.container.style.fontFamily = "Segoe UI, sans-serif";

        if (!dataView?.table) {
            this.showMessage("Add TS Number and Values fields");
            return;
        }

        this.parseSettings(dataView);

        const table = dataView.table;
        const columns = this.extractColumns(table);
        const rows = table.rows as any[][];

        if (rows.length === 0) {
            this.showMessage("No data to display");
            return;
        }

        this.renderTable(columns, rows, width, height);
    }

    private extractColumns(table: powerbi.DataViewTable): ColumnInfo[] {
        const columns: ColumnInfo[] = [];

        table.columns.forEach((col, idx) => {
            const name = col.displayName || `Column ${idx + 1}`;
            const isNumeric = col.type?.numeric || false;

            let total = 0;
            if (isNumeric) {
                (table.rows as any[][]).forEach(row => {
                    const val = row[idx];
                    if (typeof val === "number") {
                        total += val;
                    }
                });
            }

            columns.push({ name, index: idx, isNumeric, total });
        });

        return columns;
    }

    private renderTable(columns: ColumnInfo[], rows: any[][], width: number, height: number): void {
        const s = this.settings;

        const wrapper = document.createElement("div");
        wrapper.style.display = "flex";
        wrapper.style.flexDirection = "column";
        wrapper.style.height = "100%";

        // Summary cards
        if (s.showCards) {
            const cardsDiv = this.createSummaryCards(columns, rows);
            wrapper.appendChild(cardsDiv);
        }

        // Table container
        const tableContainer = document.createElement("div");
        tableContainer.style.flex = "1";
        tableContainer.style.overflow = "auto";

        const table = document.createElement("table");
        table.style.width = "100%";
        table.style.borderCollapse = "collapse";
        table.style.fontSize = s.fontSize + "px";

        // Header
        const thead = document.createElement("thead");
        const headerRow = document.createElement("tr");
        headerRow.style.backgroundColor = s.headerColor;
        headerRow.style.color = s.headerTextColor;
        headerRow.style.position = "sticky";
        headerRow.style.top = "0";
        headerRow.style.zIndex = "10";

        columns.forEach((col, idx) => {
            const th = document.createElement("th");
            th.style.padding = "10px 8px";
            th.style.textAlign = col.isNumeric ? "right" : "left";
            th.style.cursor = "pointer";
            th.style.userSelect = "none";
            th.style.borderBottom = "2px solid #cbd5e0";
            th.style.whiteSpace = "nowrap";

            let sortIndicator = "";
            if (this.sortColumn === idx) {
                sortIndicator = this.sortAscending ? " ▲" : " ▼";
            }
            th.textContent = col.name + sortIndicator;

            th.addEventListener("click", () => {
                if (this.sortColumn === idx) {
                    this.sortAscending = !this.sortAscending;
                } else {
                    this.sortColumn = idx;
                    this.sortAscending = true;
                }
                this.host.refreshHostData();
            });

            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);
        table.appendChild(thead);

        // Body
        const tbody = document.createElement("tbody");

        let sortedRows = [...rows];
        if (this.sortColumn >= 0) {
            sortedRows.sort((a, b) => {
                const valA = a[this.sortColumn];
                const valB = b[this.sortColumn];

                if (typeof valA === "number" && typeof valB === "number") {
                    return this.sortAscending ? valA - valB : valB - valA;
                }

                const strA = String(valA || "");
                const strB = String(valB || "");
                return this.sortAscending ? strA.localeCompare(strB) : strB.localeCompare(strA);
            });
        }

        sortedRows.forEach((row, rowIdx) => {
            const tr = document.createElement("tr");
            tr.style.backgroundColor = rowIdx % 2 === 0 ? s.rowColor1 : s.rowColor2;
            tr.style.color = s.textColor;
            tr.style.height = s.rowHeight + "px";

            tr.addEventListener("mouseenter", () => {
                tr.style.backgroundColor = "#e2e8f0";
            });
            tr.addEventListener("mouseleave", () => {
                tr.style.backgroundColor = rowIdx % 2 === 0 ? s.rowColor1 : s.rowColor2;
            });

            columns.forEach((col, colIdx) => {
                const td = document.createElement("td");
                td.style.padding = "8px";
                td.style.borderBottom = "1px solid #e2e8f0";
                td.style.textAlign = col.isNumeric ? "right" : "left";

                if (colIdx === 0) {
                    td.style.fontWeight = "600";
                }

                const value = row[colIdx];

                if (typeof value === "number") {
                    td.textContent = this.formatNumber(value);
                } else {
                    td.textContent = value != null ? String(value) : "";
                }

                tr.appendChild(td);
            });

            tbody.appendChild(tr);
        });

        table.appendChild(tbody);
        tableContainer.appendChild(table);
        wrapper.appendChild(tableContainer);
        this.container.appendChild(wrapper);
    }

    private createSummaryCards(columns: ColumnInfo[], rows: any[][]): HTMLElement {
        const s = this.settings;
        const cardsDiv = document.createElement("div");
        cardsDiv.style.display = "flex";
        cardsDiv.style.gap = "15px";
        cardsDiv.style.padding = "15px";
        cardsDiv.style.flexWrap = "wrap";
        cardsDiv.style.justifyContent = "flex-end";
        cardsDiv.style.backgroundColor = s.rowColor1;
        cardsDiv.style.borderBottom = "2px solid #cbd5e0";

        // Create cards for numeric columns
        columns.forEach((col, idx) => {
            if (col.isNumeric && idx > 0) {
                const card = document.createElement("div");
                card.style.backgroundColor = s.cardBackgroundColor;
                card.style.color = s.cardTextColor;
                card.style.padding = "10px 20px";
                card.style.borderRadius = "4px";
                card.style.textAlign = "center";
                card.style.minWidth = "120px";

                const title = document.createElement("div");
                title.style.fontSize = "10px";
                title.style.opacity = "0.8";
                title.style.marginBottom = "5px";
                title.textContent = col.name;

                const value = document.createElement("div");
                value.style.fontSize = "18px";
                value.style.fontWeight = "bold";
                value.textContent = this.formatNumber(col.total);

                card.appendChild(title);
                card.appendChild(value);
                cardsDiv.appendChild(card);
            }
        });

        // Total TS count card
        const countCard = document.createElement("div");
        countCard.style.backgroundColor = s.cardBackgroundColor;
        countCard.style.color = s.cardTextColor;
        countCard.style.padding = "10px 20px";
        countCard.style.borderRadius = "4px";
        countCard.style.textAlign = "center";
        countCard.style.minWidth = "120px";

        const countTitle = document.createElement("div");
        countTitle.style.fontSize = "10px";
        countTitle.style.opacity = "0.8";
        countTitle.style.marginBottom = "5px";
        countTitle.textContent = "Total TS";

        const countValue = document.createElement("div");
        countValue.style.fontSize = "18px";
        countValue.style.fontWeight = "bold";
        countValue.textContent = rows.length.toString();

        countCard.appendChild(countTitle);
        countCard.appendChild(countValue);
        cardsDiv.appendChild(countCard);

        return cardsDiv;
    }

    private formatNumber(value: number): string {
        if (Math.abs(value) >= 1000000) {
            return (value / 1000000).toFixed(1) + "M";
        }
        if (Math.abs(value) >= 1000) {
            return (value / 1000).toFixed(1) + "K";
        }
        return Math.round(value).toLocaleString();
    }

    private showMessage(msg: string): void {
        const div = document.createElement("div");
        div.style.display = "flex";
        div.style.justifyContent = "center";
        div.style.alignItems = "center";
        div.style.height = "100%";
        div.style.color = "#888";
        div.style.fontSize = "14px";
        div.textContent = msg;
        this.container.appendChild(div);
    }

    public enumerateObjectInstances(options: powerbi.EnumerateVisualObjectInstancesOptions): powerbi.VisualObjectInstanceEnumeration {
        const s = this.settings;
        switch (options.objectName) {
            case "tableSettings":
                return [{
                    objectName: options.objectName,
                    properties: {
                        headerColor: { solid: { color: s.headerColor } },
                        rowColor1: { solid: { color: s.rowColor1 } },
                        rowColor2: { solid: { color: s.rowColor2 } },
                        fontSize: s.fontSize
                    },
                    selector: null as any
                }];
            case "summaryCards":
                return [{
                    objectName: options.objectName,
                    properties: {
                        showCards: s.showCards
                    },
                    selector: null as any
                }];
            default:
                return [];
        }
    }
}
