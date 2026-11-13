const altaGamaData = {
  "meta": {
    "brand": "Garden of the Andes",
    "distributor": "Alta Gama",
    "metric": "Unidades vendidas (sell-through en unidades)",
    "source": "AltaGama_GoA_Aug-Oct_2026.xlsx",
    "sources": [
      {
        "file": "GOA_CM Sell Through Data_2026_YTD.xlsx (histórico)",
        "periods": [
          "2026-01",
          "2026-02",
          "2026-03",
          "2026-04",
          "2026-05",
          "2026-06"
        ]
      },
      {
        "file": "GetUpImorts_GoA_Kombuchacha_July-October_10.26.26.xlsx",
        "periods": [
          "2026-07"
        ]
      },
      {
        "file": "AltaGama_GoA_Aug-Oct_2026.xlsx",
        "periods": [
          "2026-08",
          "2026-09",
          "2026-10"
        ]
      }
    ],
    "periodStart": "2026-01-01",
    "periodEnd": "2026-10-31",
    "periodLabel": "Enero 2026 – Octubre 2026",
    "cutoffDate": "2026-10-31",
    "partialPeriods": [],
    "partialNote": null,
    "notes": [
      "Alta Gama proporciona información de sell-through en unidades. El reporte no incluye precios ni valores monetarios.",
      "Todos los periodos del reporte son meses completos."
    ],
    "validation": {
      "totalRowPresent": false,
      "allMatch": true,
      "checks": [
        {
          "key": "2026-01",
          "label": "Enero 2026",
          "source": "GOA_CM Sell Through Data_2026_YTD.xlsx (histórico)",
          "excelTotal": 330,
          "computed": 330,
          "ok": true
        },
        {
          "key": "2026-02",
          "label": "Febrero 2026",
          "source": "GOA_CM Sell Through Data_2026_YTD.xlsx (histórico)",
          "excelTotal": 748,
          "computed": 748,
          "ok": true
        },
        {
          "key": "2026-03",
          "label": "Marzo 2026",
          "source": "GOA_CM Sell Through Data_2026_YTD.xlsx (histórico)",
          "excelTotal": 474,
          "computed": 474,
          "ok": true
        },
        {
          "key": "2026-04",
          "label": "Abril 2026",
          "source": "GOA_CM Sell Through Data_2026_YTD.xlsx (histórico)",
          "excelTotal": 456,
          "computed": 456,
          "ok": true
        },
        {
          "key": "2026-05",
          "label": "Mayo 2026",
          "source": "GOA_CM Sell Through Data_2026_YTD.xlsx (histórico)",
          "excelTotal": 324,
          "computed": 324,
          "ok": true
        },
        {
          "key": "2026-06",
          "label": "Junio 2026",
          "source": "GOA_CM Sell Through Data_2026_YTD.xlsx (histórico)",
          "excelTotal": 264,
          "computed": 264,
          "ok": true
        },
        {
          "key": "2026-07",
          "label": "Julio 2026",
          "source": "GetUpImorts_GoA_Kombuchacha_July-October_10.26.26.xlsx",
          "excelTotal": null,
          "computed": null,
          "ok": null
        },
        {
          "key": "2026-08",
          "label": "Agosto 2026",
          "source": "AltaGama_GoA_Aug-Oct_2026.xlsx",
          "excelTotal": 324,
          "computed": 324,
          "ok": true
        },
        {
          "key": "2026-09",
          "label": "Septiembre 2026",
          "source": "AltaGama_GoA_Aug-Oct_2026.xlsx",
          "excelTotal": 486,
          "computed": 486,
          "ok": true
        },
        {
          "key": "2026-10",
          "label": "Octubre 2026",
          "source": "AltaGama_GoA_Aug-Oct_2026.xlsx",
          "excelTotal": 378,
          "computed": 378,
          "ok": true
        }
      ],
      "overrides": [
        {
          "key": "2026-07",
          "from": "GOA_CM Sell Through Data_2026_YTD.xlsx (histórico)",
          "to": "GetUpImorts_GoA_Kombuchacha_July-October_10.26.26.xlsx",
          "before": 168,
          "after": 270,
          "partialBefore": true,
          "partialAfter": false
        },
        {
          "key": "2026-08",
          "from": "GetUpImorts_GoA_Kombuchacha_July-October_10.26.26.xlsx",
          "to": "AltaGama_GoA_Aug-Oct_2026.xlsx",
          "before": 317,
          "after": 324,
          "partialBefore": false,
          "partialAfter": false
        },
        {
          "key": "2026-09",
          "from": "GetUpImorts_GoA_Kombuchacha_July-October_10.26.26.xlsx",
          "to": "AltaGama_GoA_Aug-Oct_2026.xlsx",
          "before": 474,
          "after": 486,
          "partialBefore": false,
          "partialAfter": false
        },
        {
          "key": "2026-10",
          "from": "GetUpImorts_GoA_Kombuchacha_July-October_10.26.26.xlsx",
          "to": "AltaGama_GoA_Aug-Oct_2026.xlsx",
          "before": 294,
          "after": 378,
          "partialBefore": true,
          "partialAfter": false
        }
      ],
      "issues": [
        "AltaGama_GoA_Aug-Oct_2026.xlsx · 2026-08: GA Chai Tea 1.4oz reexpresado de 29 a 30",
        "AltaGama_GoA_Aug-Oct_2026.xlsx · 2026-08: GA Ginger Lemongrass 1.4oz reexpresado de 60 a 66",
        "AltaGama_GoA_Aug-Oct_2026.xlsx · 2026-09: GA Green Tea 1.4oz reexpresado de 96 a 102",
        "AltaGama_GoA_Aug-Oct_2026.xlsx · 2026-09: GA Pure Chamomile 0.8oz reexpresado de 66 a 72"
      ]
    }
  },
  "totals": {
    "units": 4054,
    "unitsFullMonths": 4054,
    "products": 7,
    "periods": 10,
    "fullMonths": 10,
    "partialPeriods": 0,
    "avgPeriod": 405.4,
    "avgFullMonth": 405.4
  },
  "highlights": {
    "topProduct": {
      "sku": "GA Green Tea 1.4oz",
      "shortName": "Green Tea 1.4oz",
      "units": 912,
      "share": 22.5
    },
    "bestMonth": {
      "key": "2026-02",
      "label": "Febrero 2026",
      "units": 748
    },
    "lastMonth": {
      "key": "2026-10",
      "label": "Octubre 2026",
      "units": 378,
      "delta": -108,
      "deltaPct": -22.2
    },
    "partial": null
  },
  "periods": [
    {
      "key": "2026-01",
      "label": "Enero 2026",
      "short": "Ene",
      "year": 2026,
      "month": 1,
      "start": "2026-01-01",
      "end": "2026-01-31",
      "partial": false,
      "coverageDays": 31,
      "monthDays": 31,
      "units": 330,
      "delta": null,
      "deltaPct": null
    },
    {
      "key": "2026-02",
      "label": "Febrero 2026",
      "short": "Feb",
      "year": 2026,
      "month": 2,
      "start": "2026-02-01",
      "end": "2026-02-28",
      "partial": false,
      "coverageDays": 28,
      "monthDays": 28,
      "units": 748,
      "delta": 418,
      "deltaPct": 126.7
    },
    {
      "key": "2026-03",
      "label": "Marzo 2026",
      "short": "Mar",
      "year": 2026,
      "month": 3,
      "start": "2026-03-01",
      "end": "2026-03-31",
      "partial": false,
      "coverageDays": 31,
      "monthDays": 31,
      "units": 474,
      "delta": -274,
      "deltaPct": -36.6
    },
    {
      "key": "2026-04",
      "label": "Abril 2026",
      "short": "Abr",
      "year": 2026,
      "month": 4,
      "start": "2026-04-01",
      "end": "2026-04-30",
      "partial": false,
      "coverageDays": 30,
      "monthDays": 30,
      "units": 456,
      "delta": -18,
      "deltaPct": -3.8
    },
    {
      "key": "2026-05",
      "label": "Mayo 2026",
      "short": "May",
      "year": 2026,
      "month": 5,
      "start": "2026-05-01",
      "end": "2026-05-31",
      "partial": false,
      "coverageDays": 31,
      "monthDays": 31,
      "units": 324,
      "delta": -132,
      "deltaPct": -28.9
    },
    {
      "key": "2026-06",
      "label": "Junio 2026",
      "short": "Jun",
      "year": 2026,
      "month": 6,
      "start": "2026-06-01",
      "end": "2026-06-30",
      "partial": false,
      "coverageDays": 30,
      "monthDays": 30,
      "units": 264,
      "delta": -60,
      "deltaPct": -18.5
    },
    {
      "key": "2026-07",
      "label": "Julio 2026",
      "short": "Jul",
      "year": 2026,
      "month": 7,
      "start": "2026-07-01",
      "end": "2026-07-31",
      "partial": false,
      "coverageDays": 31,
      "monthDays": 31,
      "units": 270,
      "delta": 6,
      "deltaPct": 2.3
    },
    {
      "key": "2026-08",
      "label": "Agosto 2026",
      "short": "Ago",
      "year": 2026,
      "month": 8,
      "start": "2026-08-01",
      "end": "2026-08-31",
      "partial": false,
      "coverageDays": 31,
      "monthDays": 31,
      "units": 324,
      "delta": 54,
      "deltaPct": 20.0
    },
    {
      "key": "2026-09",
      "label": "Septiembre 2026",
      "short": "Sep",
      "year": 2026,
      "month": 9,
      "start": "2026-09-01",
      "end": "2026-09-30",
      "partial": false,
      "coverageDays": 30,
      "monthDays": 30,
      "units": 486,
      "delta": 162,
      "deltaPct": 50.0
    },
    {
      "key": "2026-10",
      "label": "Octubre 2026",
      "short": "Oct",
      "year": 2026,
      "month": 10,
      "start": "2026-10-01",
      "end": "2026-10-31",
      "partial": false,
      "coverageDays": 31,
      "monthDays": 31,
      "units": 378,
      "delta": -108,
      "deltaPct": -22.2
    }
  ],
  "products": [
    {
      "sku": "GA Green Tea 1.4oz",
      "name": "Garden of the Andes - Green Tea 1.4oz",
      "shortName": "Green Tea 1.4oz",
      "byPeriod": {
        "2026-01": 138,
        "2026-02": 144,
        "2026-03": 72,
        "2026-04": 114,
        "2026-05": 72,
        "2026-06": 66,
        "2026-07": 72,
        "2026-08": 72,
        "2026-09": 102,
        "2026-10": 60
      },
      "units": 912,
      "unitsFullMonths": 912,
      "share": 22.5,
      "avgFullMonth": 91.2,
      "missingPeriods": [],
      "best": {
        "key": "2026-02",
        "label": "Febrero 2026",
        "units": 144
      }
    },
    {
      "sku": "GA Roseship & Hibiscus 1.8oz",
      "name": "Garden of the Andes - Roseship and Hibiscus 1.8oz",
      "shortName": "Roseship and Hibiscus 1.8oz",
      "byPeriod": {
        "2026-01": 42,
        "2026-02": 102,
        "2026-03": 132,
        "2026-04": 66,
        "2026-05": 36,
        "2026-06": 84,
        "2026-07": 96,
        "2026-08": 72,
        "2026-09": 84,
        "2026-10": 36
      },
      "units": 750,
      "unitsFullMonths": 750,
      "share": 18.5,
      "avgFullMonth": 75.0,
      "missingPeriods": [],
      "best": {
        "key": "2026-03",
        "label": "Marzo 2026",
        "units": 132
      }
    },
    {
      "sku": "GA Ginger Lemongrass 1.4oz",
      "name": "Garden of the Andes - Ginger Lemongrass 1.4oz",
      "shortName": "Ginger Lemongrass 1.4oz",
      "byPeriod": {
        "2026-01": 108,
        "2026-02": 84,
        "2026-03": 114,
        "2026-04": 90,
        "2026-05": 66,
        "2026-06": 0,
        "2026-07": 42,
        "2026-08": 66,
        "2026-09": 78,
        "2026-10": 60
      },
      "units": 708,
      "unitsFullMonths": 708,
      "share": 17.5,
      "avgFullMonth": 70.8,
      "missingPeriods": [],
      "best": {
        "key": "2026-03",
        "label": "Marzo 2026",
        "units": 114
      }
    },
    {
      "sku": "GA Pure Peppermint 1oz",
      "name": "Garden of the Andes - Pure Peppermint 1oz",
      "shortName": "Pure Peppermint 1oz",
      "byPeriod": {
        "2026-01": 12,
        "2026-02": 126,
        "2026-03": 84,
        "2026-04": 84,
        "2026-05": 78,
        "2026-06": 48,
        "2026-07": 36,
        "2026-08": 30,
        "2026-09": 84,
        "2026-10": 102
      },
      "units": 684,
      "unitsFullMonths": 684,
      "share": 16.9,
      "avgFullMonth": 68.4,
      "missingPeriods": [],
      "best": {
        "key": "2026-02",
        "label": "Febrero 2026",
        "units": 126
      }
    },
    {
      "sku": "GA Pure Chamomile 0.8oz",
      "name": "Garden of the Andes - Pure Chamomile 0.8oz",
      "shortName": "Pure Chamomile 0.8oz",
      "byPeriod": {
        "2026-01": 12,
        "2026-02": 142,
        "2026-03": 66,
        "2026-04": 90,
        "2026-05": 66,
        "2026-06": 42,
        "2026-07": 12,
        "2026-08": 42,
        "2026-09": 72,
        "2026-10": 90
      },
      "units": 634,
      "unitsFullMonths": 634,
      "share": 15.6,
      "avgFullMonth": 63.4,
      "missingPeriods": [],
      "best": {
        "key": "2026-02",
        "label": "Febrero 2026",
        "units": 142
      }
    },
    {
      "sku": "GA Chai Tea 1.4oz",
      "name": "Garden of the Andes - Chai Tea 1.4oz",
      "shortName": "Chai Tea 1.4oz",
      "byPeriod": {
        "2026-01": 6,
        "2026-02": 114,
        "2026-03": 6,
        "2026-04": 12,
        "2026-05": 6,
        "2026-06": 24,
        "2026-07": 12,
        "2026-08": 30,
        "2026-09": 36,
        "2026-10": 30
      },
      "units": 276,
      "unitsFullMonths": 276,
      "share": 6.8,
      "avgFullMonth": 27.6,
      "missingPeriods": [],
      "best": {
        "key": "2026-02",
        "label": "Febrero 2026",
        "units": 114
      }
    },
    {
      "sku": "GA Pure Ceylon Tea 1.4oz",
      "name": "Garden of the Andes - Pure Ceylon Tea 1.4oz",
      "shortName": "Pure Ceylon Tea 1.4oz",
      "byPeriod": {
        "2026-01": 12,
        "2026-02": 36,
        "2026-03": 0,
        "2026-04": 0,
        "2026-05": 0,
        "2026-06": 0,
        "2026-07": 0,
        "2026-08": 12,
        "2026-09": 30,
        "2026-10": 0
      },
      "units": 90,
      "unitsFullMonths": 90,
      "share": 2.2,
      "avgFullMonth": 9.0,
      "missingPeriods": [],
      "best": {
        "key": "2026-02",
        "label": "Febrero 2026",
        "units": 36
      }
    }
  ]
};
