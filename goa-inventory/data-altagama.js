const altaGamaData = {
  "meta": {
    "brand": "Garden of the Andes",
    "distributor": "Alta Gama",
    "metric": "Unidades vendidas (sell-through en unidades)",
    "source": "GOA_CM Sell Through Data_2026_YTD.xlsx",
    "periodStart": "2026-01-01",
    "periodEnd": "2026-07-25",
    "periodLabel": "Enero 2026 – Julio 2026",
    "cutoffDate": "2026-07-25",
    "partialPeriods": [
      "2026-07"
    ],
    "partialNote": {
      "key": "2026-07",
      "short": "Julio 2026 incluye ventas hasta el 25 de julio. Los demás meses están completos.",
      "note": "Los datos de julio corresponden al periodo comprendido entre el 1 y el 25 de julio de 2026."
    },
    "notes": [
      "Alta Gama proporciona información de sell-through en unidades. El reporte no incluye precios ni valores monetarios.",
      "Los datos de julio corresponden al periodo comprendido entre el 1 y el 25 de julio de 2026."
    ],
    "validation": {
      "totalRowPresent": true,
      "allMatch": true,
      "checks": [
        {
          "key": "2026-01",
          "label": "Enero 2026",
          "excelTotal": 330,
          "computed": 330,
          "ok": true
        },
        {
          "key": "2026-02",
          "label": "Febrero 2026",
          "excelTotal": 748,
          "computed": 748,
          "ok": true
        },
        {
          "key": "2026-03",
          "label": "Marzo 2026",
          "excelTotal": 474,
          "computed": 474,
          "ok": true
        },
        {
          "key": "2026-04",
          "label": "Abril 2026",
          "excelTotal": 456,
          "computed": 456,
          "ok": true
        },
        {
          "key": "2026-05",
          "label": "Mayo 2026",
          "excelTotal": 324,
          "computed": 324,
          "ok": true
        },
        {
          "key": "2026-06",
          "label": "Junio 2026",
          "excelTotal": 264,
          "computed": 264,
          "ok": true
        },
        {
          "key": "2026-07",
          "label": "Julio 2026",
          "excelTotal": 168,
          "computed": 168,
          "ok": true
        }
      ],
      "issues": []
    }
  },
  "totals": {
    "units": 2764,
    "unitsFullMonths": 2596,
    "products": 7,
    "periods": 7,
    "fullMonths": 6,
    "partialPeriods": 1,
    "avgPeriod": 394.9,
    "avgFullMonth": 432.7
  },
  "highlights": {
    "topProduct": {
      "sku": "GA Green Tea 1.4oz",
      "shortName": "Green Tea 1.4oz",
      "units": 660,
      "share": 23.9
    },
    "bestMonth": {
      "key": "2026-02",
      "label": "Febrero 2026",
      "units": 748
    },
    "lastMonth": {
      "key": "2026-07",
      "label": "Julio 2026",
      "units": 168,
      "delta": -96,
      "deltaPct": -36.4
    },
    "partial": {
      "key": "2026-07",
      "label": "Julio 2026",
      "short": "Jul",
      "units": 168,
      "start": "2026-07-01",
      "end": "2026-07-25",
      "coverageDays": 25,
      "monthDays": 31
    }
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
      "end": "2026-07-25",
      "partial": true,
      "coverageDays": 25,
      "monthDays": 31,
      "units": 168,
      "delta": -96,
      "deltaPct": -36.4
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
        "2026-07": 54
      },
      "units": 660,
      "unitsFullMonths": 606,
      "share": 23.9,
      "avgFullMonth": 101.0,
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
        "2026-07": 60
      },
      "units": 522,
      "unitsFullMonths": 462,
      "share": 18.9,
      "avgFullMonth": 77.0,
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
        "2026-07": 18
      },
      "units": 480,
      "unitsFullMonths": 462,
      "share": 17.4,
      "avgFullMonth": 77.0,
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
        "2026-07": 12
      },
      "units": 444,
      "unitsFullMonths": 432,
      "share": 16.1,
      "avgFullMonth": 72.0,
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
        "2026-07": 12
      },
      "units": 430,
      "unitsFullMonths": 418,
      "share": 15.6,
      "avgFullMonth": 69.7,
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
        "2026-07": 12
      },
      "units": 180,
      "unitsFullMonths": 168,
      "share": 6.5,
      "avgFullMonth": 28.0,
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
        "2026-07": 0
      },
      "units": 48,
      "unitsFullMonths": 48,
      "share": 1.7,
      "avgFullMonth": 8.0,
      "missingPeriods": [],
      "best": {
        "key": "2026-02",
        "label": "Febrero 2026",
        "units": 36
      }
    }
  ]
};
