import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  AreaChart, Area, BarChart, Bar, LineChart, Line, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend,
} from "recharts";
import { Upload, PlusCircle, Download, TrendingUp, TrendingDown, HardHat, MapPin, Clock, AlertCircle, ChevronDown, ChevronUp, X, Check, Car, Users, Trash2, ClipboardCheck, Siren, Lock, ArrowRight, ArrowLeft } from "lucide-react";

/* ---------------------------------------------------------------------- */
/* Seed data — parsed from Mentari HSE Data MASTER, last updated 2026-08-01 */
/* ---------------------------------------------------------------------- */

const SITE_CATEGORIES = ["KMS", "Survey", "Driver", "PBS", "JGC", "GSB", "KKS", "WEN", "TSE", "SOS", "PME", "YKG", "MPE", "PTTS", "SAFECON", "THERMAX (INDIA)", "BND", "PT. Siemens"];
const OFFSITE_CATEGORIES = ["JGC", "PBS", "GSB", "KKS", "WEN", "THERMAX (INDIA)", "THERMAX (INDONESIA)", "MAJU BERSAMA ", "SEFCON", "MECC2000", "LIPICO", "PM ELECTRIC", "PANWATER", "MECGALE", "PTTS", "DBTS", "PPESB", "STATRON", "PT. Siemens", "PT MUGI"];
const PROJECT_START = "2025-04";
const DATA_AS_OF = "Aug 2026";
const BASELINE_CUTOFF = "2026-07"; // indicator baseline is treated as solid through this month; Aug 2026 onward is entered live via "Log month"
const CURRENT_EDITABLE_MONTH = "2026-08"; // "Log month" opens on this month by default
const SEED_MONTHS = [{"month":"2025-01","total":7645.0,"site":{"KMS":0.0,"Survey":0.0,"Driver":0.0,"PBS":0.0,"JGC":0.0,"GSB":0.0,"KKS":0.0,"WEN":0.0,"TSE":0.0,"SOS":0.0,"PM ELECTRIC":0.0,"Yokogawa":0.0,"THERMAX (INDIA)":0.0,"PTTS":0.0,"BND":0.0,"PT. Siemens":0.0,"MPE":0.0,"Safecon":0.0},"offsite":{"JGC":6056.0,"PBS":0.0,"GSB":0.0,"KKS":0.0,"WEN":0.0,"THERMAX (INDIA)":0.0,"THERMAX (INDO)":0.0,"MAJU BERSAMA ":0.0,"SEFCON":0.0,"PTTS":0.0,"DBTS":0.0,"PANWATER":0.0,"MECGALE":0.0,"LIPICO":0.0,"MECC2000":0.0,"PPESB":0.0,"STATRON":0.0,"PT. Siemens":0.0,"PT MUGI":0.0},"staffPermanent":{"site":0.0,"offsite":1589.0},"indicators":{}},{"month":"2025-02","total":13082.0,"site":{"KMS":0.0,"Survey":0.0,"Driver":0.0,"PBS":0.0,"JGC":0.0,"GSB":0.0,"KKS":0.0,"WEN":0.0,"TSE":0.0,"SOS":0.0,"PM ELECTRIC":0.0,"Yokogawa":0.0,"THERMAX (INDIA)":0.0,"PTTS":0.0,"BND":0.0,"PT. Siemens":0.0,"MPE":0.0,"Safecon":0.0},"offsite":{"JGC":11509.0,"PBS":0.0,"GSB":0.0,"KKS":0.0,"WEN":0.0,"THERMAX (INDIA)":0.0,"THERMAX (INDO)":0.0,"MAJU BERSAMA ":0.0,"SEFCON":0.0,"PTTS":0.0,"DBTS":0.0,"PANWATER":0.0,"MECGALE":0.0,"LIPICO":0.0,"MECC2000":0.0,"PPESB":0.0,"STATRON":0.0,"PT. Siemens":0.0,"PT MUGI":0.0},"staffPermanent":{"site":0.0,"offsite":1573.0},"indicators":{}},{"month":"2025-03","total":14586.0,"site":{"KMS":0.0,"Survey":0.0,"Driver":0.0,"PBS":0.0,"JGC":0.0,"GSB":0.0,"KKS":0.0,"WEN":0.0,"TSE":0.0,"SOS":0.0,"PM ELECTRIC":0.0,"Yokogawa":0.0,"THERMAX (INDIA)":0.0,"PTTS":0.0,"BND":0.0,"PT. Siemens":0.0,"MPE":0.0,"Safecon":0.0},"offsite":{"JGC":12913.0,"PBS":0.0,"GSB":0.0,"KKS":0.0,"WEN":0.0,"THERMAX (INDIA)":0.0,"THERMAX (INDO)":0.0,"MAJU BERSAMA ":0.0,"SEFCON":0.0,"PTTS":0.0,"DBTS":0.0,"PANWATER":0.0,"MECGALE":0.0,"LIPICO":0.0,"MECC2000":0.0,"PPESB":0.0,"STATRON":0.0,"PT. Siemens":0.0,"PT MUGI":0.0},"staffPermanent":{"site":0.0,"offsite":1673.0},"indicators":{}},{"month":"2025-04","total":16981.4,"site":{"KMS":1982.4,"Survey":384.0,"Driver":120.0,"PBS":0.0,"JGC":0.0,"GSB":0.0,"KKS":0.0,"WEN":0.0,"TSE":0.0,"SOS":0.0,"PM ELECTRIC":0.0,"Yokogawa":0.0,"THERMAX (INDIA)":0.0,"PTTS":0.0,"BND":0.0,"PT. Siemens":0.0,"MPE":0.0,"Safecon":0.0},"offsite":{"JGC":12784.0,"PBS":0.0,"GSB":0.0,"KKS":0.0,"WEN":0.0,"THERMAX (INDIA)":0.0,"THERMAX (INDO)":0.0,"MAJU BERSAMA ":0.0,"SEFCON":0.0,"PTTS":0.0,"DBTS":0.0,"PANWATER":0.0,"MECGALE":0.0,"LIPICO":0.0,"MECC2000":0.0,"PPESB":0.0,"STATRON":0.0,"PT. Siemens":0.0,"PT MUGI":0.0},"staffPermanent":{"site":242.0,"offsite":1469.0},"indicators":{}},{"month":"2025-05","total":26138.4,"site":{"KMS":4664.4,"Survey":885.0,"Driver":336.0,"PBS":3941.0,"JGC":0.0,"GSB":0.0,"KKS":0.0,"WEN":0.0,"TSE":0.0,"SOS":0.0,"PM ELECTRIC":0.0,"Yokogawa":0.0,"THERMAX (INDIA)":0.0,"PTTS":0.0,"BND":0.0,"PT. Siemens":0.0,"MPE":0.0,"Safecon":0.0},"offsite":{"JGC":14257.0,"PBS":0.0,"GSB":0.0,"KKS":0.0,"WEN":0.0,"THERMAX (INDIA)":0.0,"THERMAX (INDO)":0.0,"MAJU BERSAMA ":0.0,"SEFCON":0.0,"PTTS":0.0,"DBTS":0.0,"PANWATER":0.0,"MECGALE":0.0,"LIPICO":0.0,"MECC2000":0.0,"PPESB":0.0,"STATRON":0.0,"PT. Siemens":0.0,"PT MUGI":0.0},"staffPermanent":{"site":486.0,"offsite":1569.0},"indicators":{}},{"month":"2025-06","total":53188.856,"site":{"KMS":4282.4,"Survey":957.0,"Driver":369.0,"PBS":29739.0,"JGC":280.0,"GSB":0.0,"KKS":0.0,"WEN":0.0,"TSE":0.0,"SOS":0.0,"PM ELECTRIC":0.0,"Yokogawa":0.0,"THERMAX (INDIA)":0.0,"PTTS":0.0,"BND":0.0,"PT. Siemens":0.0,"MPE":0.0,"Safecon":0.0},"offsite":{"JGC":14923.0,"PBS":0.0,"GSB":0.0,"KKS":0.0,"WEN":0.0,"THERMAX (INDIA)":0.0,"THERMAX (INDO)":0.0,"MAJU BERSAMA ":579.4559999999999,"SEFCON":0.0,"PTTS":0.0,"DBTS":0.0,"PANWATER":0.0,"MECGALE":0.0,"LIPICO":0.0,"MECC2000":0.0,"PPESB":0.0,"STATRON":0.0,"PT. Siemens":0.0,"PT MUGI":0.0},"staffPermanent":{"site":503.0,"offsite":1556.0},"indicators":{}},{"month":"2025-07","total":67221.809,"site":{"KMS":1586.4,"Survey":972.0,"Driver":648.0,"PBS":39741.0,"JGC":400.0,"GSB":0.0,"KKS":0.0,"WEN":0.0,"TSE":0.0,"SOS":0.0,"PM ELECTRIC":0.0,"Yokogawa":0.0,"THERMAX (INDIA)":0.0,"PTTS":0.0,"BND":0.0,"PT. Siemens":0.0,"MPE":0.0,"Safecon":0.0},"offsite":{"JGC":18668.0,"PBS":0.0,"GSB":0.0,"KKS":0.0,"WEN":0.0,"THERMAX (INDIA)":0.0,"THERMAX (INDO)":0.0,"MAJU BERSAMA ":2800.4090000000006,"SEFCON":0.0,"PTTS":0.0,"DBTS":0.0,"PANWATER":0.0,"MECGALE":0.0,"LIPICO":0.0,"MECC2000":0.0,"PPESB":0.0,"STATRON":0.0,"PT. Siemens":0.0,"PT MUGI":0.0},"staffPermanent":{"site":467.0,"offsite":1939.0},"indicators":{}},{"month":"2025-08","total":86135.731,"site":{"KMS":562.4,"Survey":2124.0,"Driver":936.0,"PBS":52496.0,"JGC":360.0,"GSB":0.0,"KKS":0.0,"WEN":0.0,"TSE":0.0,"SOS":0.0,"PM ELECTRIC":0.0,"Yokogawa":0.0,"THERMAX (INDIA)":0.0,"PTTS":0.0,"BND":0.0,"PT. Siemens":0.0,"MPE":0.0,"Safecon":0.0},"offsite":{"JGC":16277.0,"PBS":0.0,"GSB":7680.0,"KKS":0.0,"WEN":0.0,"THERMAX (INDIA)":0.0,"THERMAX (INDO)":0.0,"MAJU BERSAMA ":2233.331,"SEFCON":0.0,"PTTS":0.0,"DBTS":0.0,"PANWATER":0.0,"MECGALE":0.0,"LIPICO":0.0,"MECC2000":0.0,"PPESB":0.0,"STATRON":0.0,"PT. Siemens":0.0,"PT MUGI":0.0},"staffPermanent":{"site":1528.0,"offsite":1939.0},"indicators":{}},{"month":"2025-09","total":96699.788,"site":{"KMS":0.0,"Survey":2007.0,"Driver":1260.0,"PBS":49988.0,"JGC":300.0,"GSB":7336.0,"KKS":5270.0,"WEN":0.0,"TSE":0.0,"SOS":0.0,"PM ELECTRIC":0.0,"Yokogawa":0.0,"THERMAX (INDIA)":0.0,"PTTS":0.0,"BND":0.0,"PT. Siemens":0.0,"MPE":0.0,"Safecon":0.0},"offsite":{"JGC":16458.0,"PBS":0.0,"GSB":6144.0,"KKS":0.0,"WEN":0.0,"THERMAX (INDIA)":0.0,"THERMAX (INDO)":0.0,"MAJU BERSAMA ":2894.7880000000005,"SEFCON":0.0,"PTTS":0.0,"DBTS":0.0,"PANWATER":0.0,"MECGALE":0.0,"LIPICO":0.0,"MECC2000":0.0,"PPESB":0.0,"STATRON":0.0,"PT. Siemens":0.0,"PT MUGI":0.0},"staffPermanent":{"site":3103.0,"offsite":1939.0},"indicators":{}},{"month":"2025-10","total":127771.221,"site":{"KMS":0.0,"Survey":1728.0,"Driver":936.0,"PBS":45202.0,"JGC":440.0,"GSB":9568.0,"KKS":25077.0,"WEN":0.0,"TSE":0.0,"SOS":0.0,"PM ELECTRIC":0.0,"Yokogawa":0.0,"THERMAX (INDIA)":0.0,"PTTS":0.0,"BND":0.0,"PT. Siemens":0.0,"MPE":0.0,"Safecon":0.0},"offsite":{"JGC":18836.0,"PBS":11504.0,"GSB":6144.0,"KKS":0.0,"WEN":0.0,"THERMAX (INDIA)":0.0,"THERMAX (INDO)":0.0,"MAJU BERSAMA ":3784.220999999998,"SEFCON":0.0,"PTTS":0.0,"DBTS":0.0,"PANWATER":0.0,"MECGALE":0.0,"LIPICO":0.0,"MECC2000":0.0,"PPESB":0.0,"STATRON":0.0,"PT. Siemens":0.0,"PT MUGI":0.0},"staffPermanent":{"site":2789.0,"offsite":1763.0},"indicators":{}},{"month":"2025-11","total":125741.951,"site":{"KMS":0.0,"Survey":2016.0,"Driver":1008.0,"PBS":46734.0,"JGC":600.0,"GSB":11616.0,"KKS":24081.0,"WEN":0.0,"TSE":0.0,"SOS":0.0,"PM ELECTRIC":0.0,"Yokogawa":0.0,"THERMAX (INDIA)":0.0,"PTTS":0.0,"BND":0.0,"PT. Siemens":0.0,"MPE":0.0,"Safecon":0.0},"offsite":{"JGC":15680.0,"PBS":9950.0,"GSB":7680.0,"KKS":0.0,"WEN":0.0,"THERMAX (INDIA)":0.0,"THERMAX (INDO)":0.0,"MAJU BERSAMA ":1615.951000000001,"SEFCON":0.0,"PTTS":0.0,"DBTS":0.0,"PANWATER":0.0,"MECGALE":0.0,"LIPICO":0.0,"MECC2000":0.0,"PPESB":0.0,"STATRON":0.0,"PT. Siemens":0.0,"PT MUGI":0.0},"staffPermanent":{"site":3174.0,"offsite":1587.0},"indicators":{}},{"month":"2025-12","total":185285.83,"site":{"KMS":0.0,"Survey":2124.0,"Driver":944.0,"PBS":63917.0,"JGC":260.0,"GSB":14552.0,"KKS":30249.0,"WEN":1621.0,"TSE":0.0,"SOS":0.0,"PM ELECTRIC":0.0,"Yokogawa":0.0,"THERMAX (INDIA)":0.0,"PTTS":0.0,"BND":0.0,"PT. Siemens":0.0,"MPE":0.0,"Safecon":0.0},"offsite":{"JGC":15099.0,"PBS":9914.0,"GSB":5888.0,"KKS":0.0,"WEN":0.0,"THERMAX (INDIA)":28500.0,"THERMAX (INDO)":0.0,"MAJU BERSAMA ":2829.8299999999995,"SEFCON":4800.0,"PTTS":0.0,"DBTS":0.0,"PANWATER":0.0,"MECGALE":0.0,"LIPICO":0.0,"MECC2000":0.0,"PPESB":0.0,"STATRON":0.0,"PT. Siemens":0.0,"PT MUGI":0.0},"staffPermanent":{"site":2869.0,"offsite":1719.0},"indicators":{}},{"month":"2026-01","total":168062.25,"site":{"KMS":0.0,"Survey":1881.0,"Driver":1136.0,"PBS":57041.0,"JGC":240.0,"GSB":16272.0,"KKS":31785.0,"WEN":8472.0,"TSE":0.0,"SOS":0.0,"PM ELECTRIC":0.0,"Yokogawa":0.0,"THERMAX (INDIA)":0.0,"PTTS":0.0,"BND":0.0,"PT. Siemens":0.0,"MPE":0.0,"Safecon":0.0},"offsite":{"JGC":13773.0,"PBS":15432.0,"GSB":10104.0,"KKS":0.0,"WEN":0.0,"THERMAX (INDIA)":1500.0,"THERMAX (INDO)":0.0,"MAJU BERSAMA ":734.0,"SEFCON":4800.0,"PTTS":0.0,"DBTS":0.0,"PANWATER":0.0,"MECGALE":0.0,"LIPICO":0.0,"MECC2000":0.0,"PPESB":0.0,"STATRON":0.0,"PT. Siemens":0.0,"PT MUGI":0.0},"staffPermanent":{"site":3339.25,"offsite":1553.0},"indicators":{}},{"month":"2026-02","total":312591.0,"site":{"KMS":0.0,"Survey":2016.0,"Driver":1000.0,"PBS":38589.0,"JGC":280.0,"GSB":18008.0,"KKS":41887.0,"WEN":14601.0,"TSE":30.0,"SOS":0.0,"PM ELECTRIC":0.0,"Yokogawa":0.0,"THERMAX (INDIA)":0.0,"PTTS":0.0,"BND":0.0,"PT. Siemens":0.0,"MPE":0.0,"Safecon":0.0},"offsite":{"JGC":8907.0,"PBS":8708.0,"GSB":5888.0,"KKS":158111.0,"WEN":0.0,"THERMAX (INDIA)":1500.0,"THERMAX (INDO)":0.0,"MAJU BERSAMA ":2531.0,"SEFCON":4800.0,"PTTS":0.0,"DBTS":0.0,"PANWATER":0.0,"MECGALE":0.0,"LIPICO":0.0,"MECC2000":0.0,"PPESB":0.0,"STATRON":0.0,"PT. Siemens":0.0,"PT MUGI":0.0},"staffPermanent":{"site":3963.0,"offsite":1772.0},"indicators":{}},{"month":"2026-03","total":160918.014,"site":{"KMS":0.0,"Survey":1809.0,"Driver":978.0,"PBS":35624.0,"JGC":230.0,"GSB":13856.0,"KKS":47612.0,"WEN":12989.0,"TSE":168.0,"SOS":0.0,"PM ELECTRIC":0.0,"Yokogawa":0.0,"THERMAX (INDIA)":0.0,"PTTS":0.0,"BND":0.0,"PT. Siemens":0.0,"MPE":0.0,"Safecon":0.0},"offsite":{"JGC":440.0,"PBS":3480.0,"GSB":4608.0,"KKS":24310.0,"WEN":0.0,"THERMAX (INDIA)":1500.0,"THERMAX (INDO)":0.0,"MAJU BERSAMA ":918.0140000000029,"SEFCON":4800.0,"PTTS":1673.0,"DBTS":550.0,"PANWATER":0.0,"MECGALE":0.0,"LIPICO":0.0,"MECC2000":0.0,"PPESB":0.0,"STATRON":0.0,"PT. Siemens":0.0,"PT MUGI":0.0},"staffPermanent":{"site":3705.0,"offsite":1668.0},"indicators":{}},{"month":"2026-04","total":312860.0,"site":{"KMS":0.0,"Survey":2256.0,"Driver":1116.0,"PBS":74472.0,"JGC":20.0,"GSB":23592.0,"KKS":114933.0,"WEN":56087.0,"TSE":281.0,"SOS":765.0,"PM ELECTRIC":176.0,"Yokogawa":384.0,"THERMAX (INDIA)":0.0,"PTTS":0.0,"BND":0.0,"PT. Siemens":0.0,"MPE":0.0,"Safecon":0.0},"offsite":{"JGC":560.0,"PBS":0.0,"GSB":6656.0,"KKS":16916.0,"WEN":0.0,"THERMAX (INDIA)":500.0,"THERMAX (INDO)":0.0,"MAJU BERSAMA ":0.0,"SEFCON":4800.0,"PTTS":0.0,"DBTS":0.0,"PANWATER":0.0,"MECGALE":800.0,"LIPICO":0.0,"MECC2000":0.0,"PPESB":1872.0,"STATRON":0.0,"PT. Siemens":0.0,"PT MUGI":0.0},"staffPermanent":{"site":4608.0,"offsite":2066.0},"indicators":{}},{"month":"2026-05","total":304959.8,"site":{"KMS":0.0,"Survey":2160.0,"Driver":1116.0,"PBS":79110.0,"JGC":0.0,"GSB":19464.0,"KKS":123650.0,"WEN":43642.0,"TSE":135.5,"SOS":2410.0,"PM ELECTRIC":0.0,"Yokogawa":0.0,"THERMAX (INDIA)":0.0,"PTTS":0.0,"BND":0.0,"PT. Siemens":0.0,"MPE":0.0,"Safecon":0.0},"offsite":{"JGC":440.0,"PBS":0.0,"GSB":6656.0,"KKS":21036.0,"WEN":0.0,"THERMAX (INDIA)":0.0,"THERMAX (INDO)":0.0,"MAJU BERSAMA ":0.0,"SEFCON":0.0,"PTTS":0.0,"DBTS":0.0,"PANWATER":0.0,"MECGALE":0.0,"LIPICO":0.0,"MECC2000":0.0,"PPESB":0.0,"STATRON":0.0,"PT. Siemens":0.0,"PT MUGI":0.0},"staffPermanent":{"site":3753.0,"offsite":1387.3},"indicators":{}},{"month":"2026-06","total":292813.5,"site":{"KMS":0.0,"Survey":2088.0,"Driver":984.0,"PBS":80040.0,"JGC":0.0,"GSB":23104.0,"KKS":122121.0,"WEN":41765.0,"TSE":261.5,"SOS":2340.0,"PM ELECTRIC":0.0,"Yokogawa":0.0,"THERMAX (INDIA)":200.0,"PTTS":1800.0,"BND":24.0,"PT. Siemens":30.0,"MPE":0.0,"Safecon":0.0},"offsite":{"JGC":520.0,"PBS":0.0,"GSB":6656.0,"KKS":0.0,"WEN":0.0,"THERMAX (INDIA)":0.0,"THERMAX (INDO)":0.0,"MAJU BERSAMA ":0.0,"SEFCON":0.0,"PTTS":0.0,"DBTS":0.0,"PANWATER":0.0,"MECGALE":0.0,"LIPICO":5000.0,"MECC2000":0.0,"PPESB":0.0,"STATRON":0.0,"PT. Siemens":0.0,"PT MUGI":0.0},"staffPermanent":{"site":4199.0,"offsite":1681.0},"indicators":{"km":{"site":10402,"offsite":0}}},{"month":"2026-07","total":334934.0,"site":{"KMS":0.0,"Survey":2232.0,"Driver":1100.0,"PBS":85889.0,"JGC":0.0,"GSB":21478.0,"KKS":165974.0,"WEN":35814.0,"TSE":290.5,"SOS":3132.0,"PM ELECTRIC":0.0,"Yokogawa":0.0,"THERMAX (INDIA)":310.0,"PTTS":4290.0,"BND":0.0,"PT. Siemens":0.0,"MPE":275.0,"Safecon":380.0},"offsite":{"JGC":540.0,"PBS":0.0,"GSB":6912.0,"KKS":0.0,"WEN":0.0,"THERMAX (INDIA)":0.0,"THERMAX (INDO)":0.0,"MAJU BERSAMA ":0.0,"SEFCON":0.0,"PTTS":0.0,"DBTS":0.0,"PANWATER":0.0,"MECGALE":0.0,"LIPICO":0.0,"MECC2000":0.0,"PPESB":0.0,"STATRON":0.0,"PT. Siemens":0.0,"PT MUGI":0.0},"staffPermanent":{"site":4360.0,"offsite":1957.5},"indicators":{"km":{"site":10580,"offsite":0}}},{"month":"2026-08","total":288257.6,"site":{"KMS":0.0,"Survey":2160.0,"Driver":1032.0,"PBS":72643.0,"JGC":0.0,"GSB":17936.0,"KKS":142305.0,"WEN":29772.0,"TSE":304.5,"SOS":4856.0,"PM ELECTRIC":0.0,"Yokogawa":0.0,"THERMAX (INDIA)":285.0,"PTTS":3750.0,"BND":0.0,"PT. Siemens":0.0,"MPE":252.0,"Safecon":0.0},"offsite":{"JGC":500.0,"PBS":0.0,"GSB":6400.0,"KKS":0.0,"WEN":0.0,"THERMAX (INDIA)":0.0,"THERMAX (INDO)":0.0,"MAJU BERSAMA ":0.0,"SEFCON":0.0,"PTTS":0.0,"DBTS":0.0,"PANWATER":0.0,"MECGALE":0.0,"LIPICO":0.0,"MECC2000":0.0,"PPESB":0.0,"STATRON":0.0,"PT. Siemens":0.0,"PT MUGI":0.0},"staffPermanent":{"site":4294.5,"offsite":1767.6},"indicators":{"uaUc":{"site":1213,"offsite":0}}}];
const SEED_PTD_SUMMARY = {"workHours":{"ptd":2744671.55,"site":2064589.75,"offsite":680081.8},"trcf":{"ptd":0.0,"site":0.0,"offsite":0.0},"ltif":{"ptd":0.0,"site":0.0,"offsite":0.0},"lsr":{"ptd":2.0,"site":2.0,"offsite":0.0},"hipo":{"ptd":0.0,"site":0.0,"offsite":0.0},"otherSig":{"ptd":0.0,"site":0.0,"offsite":0.0},"fac":{"ptd":4.0,"site":4.0,"offsite":0.0},"oi":{"ptd":0.0,"site":0.0,"offsite":0.0},"nearMiss":{"ptd":51.0,"site":51.0,"offsite":0.0},"envIncident":{"ptd":0.0,"site":0.0,"offsite":0.0},"equipDamage":{"ptd":7.0,"site":7.0,"offsite":0.0},"security":{"ptd":7.0,"site":7.0,"offsite":0.0},"mvi":{"ptd":0.0,"site":0.0,"offsite":0.0},"transport":{"ptd":0.0,"site":0.0,"offsite":0.0},"uaUc":{"ptd":5273.0,"site":5273.0,"offsite":0.0},"mgmtVisits":{"ptd":106.0,"site":92.0,"offsite":14.0},"walkthroughs":{"ptd":294.0,"site":280.0,"offsite":14.0},"inspections":{"ptd":1745.0,"site":1689.0,"offsite":56.0},"ptw":{"ptd":3010.0,"site":2470.0,"offsite":540.0},"jha":{"ptd":349.0,"site":349.0,"offsite":0.0},"liftPlans":{"ptd":266.0,"site":266.0,"offsite":0.0},"inductions":{"ptd":2300.0,"site":2300.0,"offsite":0.0},"trainings":{"ptd":413.0,"site":413.0,"offsite":0.0},"toolboxTalks":{"ptd":3634.0,"site":3454.0,"offsite":180.0},"meetings":{"ptd":434.0,"site":394.0,"offsite":40.0},"drills":{"ptd":17.0,"site":17.0,"offsite":0.0},"audits":{"ptd":5.0,"site":5.0,"offsite":0.0},"moc":{"ptd":0.0,"site":0.0,"offsite":0.0},"standDowns":{"ptd":40.0,"site":40.0,"offsite":0.0},"drugAlcohol":{"ptd":266.0,"site":266.0,"offsite":0.0},"km":{"ptd":235072.0,"site":235072.0,"offsite":0.0},"warnings":{"ptd":214.0,"site":214.0,"offsite":0.0},"hazWaste":{"ptd":1.5,"site":1.5,"offsite":0.0},"nonHazWaste":{"ptd":132.15,"site":132.15,"offsite":0.0}};
const WEEKLY = [{"w":1,"d":"10-17 Apr","wh":570,"nm":0.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":32.0,"insp":1.0,"ptw":4.0,"tbt":7.0,"ind":0.0,"km":106.0},{"w":2,"d":"21-26 Apr","wh":1008,"nm":0.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":32.0,"insp":1.0,"ptw":4.0,"tbt":7.0,"ind":0.0,"km":106.0},{"w":3,"d":"27 Ap-03 Mayr","wh":1030,"nm":1.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":23.0,"insp":0.0,"ptw":6.0,"tbt":7.0,"ind":0.0,"km":122.0},{"w":4,"d":"04 -10 May","wh":1153,"nm":1.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":20.0,"insp":0.0,"ptw":4.0,"tbt":7.0,"ind":0.0,"km":209.0},{"w":5,"d":"11 -17 May","wh":1156,"nm":0.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":18.0,"insp":2.0,"ptw":7.0,"tbt":11.0,"ind":0.0,"km":392.0},{"w":6,"d":"18 -24 May","wh":3198,"nm":0.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":33.0,"insp":3.0,"ptw":8.0,"tbt":13.0,"ind":2.0,"km":353.0},{"w":7,"d":"25 -31 May","wh":4325,"nm":0.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":34.0,"insp":1.0,"ptw":11.0,"tbt":13.0,"ind":10.0,"km":384.0},{"w":8,"d":"01-07 -  June","wh":6020,"nm":1.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":54.0,"insp":7.0,"ptw":20.0,"tbt":12.0,"ind":38.0,"km":493.0},{"w":9,"d":"08 - 14 -  June","wh":9010,"nm":1.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":42.0,"insp":11.0,"ptw":24.0,"tbt":17.0,"ind":44.0,"km":470.0},{"w":10,"d":"15 - 21 -  June","wh":8525,"nm":0.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":42.0,"insp":2.0,"ptw":26.0,"tbt":19.0,"ind":9.0,"km":470.0},{"w":11,"d":"22 - 28 -  June","wh":10431,"nm":1.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":40.0,"insp":4.0,"ptw":23.0,"tbt":5.0,"ind":56.0,"km":1992.0},{"w":12,"d":"29 June- 05 July","wh":13392,"nm":0.0,"fac":0.0,"sec":0,"ed":1.0,"lsr":0.0,"uauc":124.0,"insp":16.0,"ptw":22.0,"tbt":6.0,"ind":37.0,"km":1612.0},{"w":13,"d":"06 - 12 July","wh":11250,"nm":0.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":120.0,"insp":9.0,"ptw":18.0,"tbt":16.0,"ind":15.0,"km":744.0},{"w":14,"d":"13 - 19 July","wh":10975,"nm":0.0,"fac":1.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":96.0,"insp":10.0,"ptw":13.0,"tbt":14.0,"ind":25.0,"km":1112.0},{"w":15,"d":"20 - 26 July","wh":9558,"nm":0.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":180.0,"insp":21.0,"ptw":13.0,"tbt":13.0,"ind":10.0,"km":4411.0},{"w":16,"d":"27 July - 02 Aug","wh":11729,"nm":0.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":161.0,"insp":14.0,"ptw":18.0,"tbt":13.0,"ind":15.0,"km":2539.0},{"w":17,"d":"03 Aug - 09 Aug","wh":11255,"nm":0.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":180.0,"insp":21.0,"ptw":13.0,"tbt":13.0,"ind":10.0,"km":4411.0},{"w":18,"d":"10 Aug - 16 Aug","wh":11189,"nm":0.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":85.0,"insp":68.0,"ptw":30.0,"tbt":13.0,"ind":10.0,"km":4411.0},{"w":19,"d":"17 Aug - 23 Aug","wh":11975,"nm":1.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":41.0,"insp":69.0,"ptw":30.0,"tbt":13.0,"ind":10.0,"km":558.0},{"w":20,"d":"24 Aug - 31Aug","wh":12001,"nm":2.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":42.0,"insp":6.0,"ptw":30.0,"tbt":13.0,"ind":388.0,"km":4411.0},{"w":21,"d":"31st Aug \u2013 06th Sept 2025","wh":10506,"nm":1.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":53.0,"insp":4.0,"ptw":30.0,"tbt":18.0,"ind":15.0,"km":4908.0},{"w":22,"d":"07- 13 Sept 2025","wh":11895,"nm":0.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":13.0,"insp":4.0,"ptw":30.0,"tbt":19.0,"ind":67.0,"km":1934.0},{"w":23,"d":"14- 20 Sept 2025","wh":14458,"nm":1.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":12.0,"insp":5.0,"ptw":30.0,"tbt":16.0,"ind":18.0,"km":1624.0},{"w":24,"d":"21- 27 Sept 2025","wh":13739,"nm":1.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":25.0,"insp":16.0,"ptw":30.0,"tbt":43.0,"ind":52.0,"km":2000.0},{"w":25,"d":"28 Sept - 04 Oct 2025","wh":13396,"nm":4.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":73.0,"insp":2.0,"ptw":30.0,"tbt":17.0,"ind":75.0,"km":3609.0},{"w":26,"d":"05th \u2013 11th Oct 2025","wh":22288,"nm":2.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":103.0,"insp":12.0,"ptw":30.0,"tbt":46.0,"ind":39.0,"km":3609.0},{"w":27,"d":"12th \u2013 18th Oct 2025","wh":20996,"nm":0.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":78.0,"insp":9.0,"ptw":30.0,"tbt":50.0,"ind":6.0,"km":3609.0},{"w":28,"d":"19th \u2013 25th Oct 2025","wh":21962,"nm":3.0,"fac":0.0,"sec":0,"ed":1.0,"lsr":0.0,"uauc":117.0,"insp":8.0,"ptw":30.0,"tbt":49.0,"ind":10.0,"km":4013.0},{"w":29,"d":"26th Oct \u2013 01 Nov 2025","wh":11381,"nm":2.0,"fac":1.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":67.0,"insp":6.0,"ptw":30.0,"tbt":30.0,"ind":5.0,"km":4013.0},{"w":30,"d":"02 - 08 - Nov 2025","wh":21152,"nm":0.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":22.0,"insp":6.0,"ptw":30.0,"tbt":51.0,"ind":9.0,"km":4013.0},{"w":31,"d":"09 Nov \u2013 15 Nov 2025","wh":22668,"nm":1.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":23.0,"insp":6.0,"ptw":30.0,"tbt":52.0,"ind":30.0,"km":4443.0},{"w":32,"d":"16 Nov \u2013 22 Nov 2025","wh":22762,"nm":2.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":23.0,"insp":6.0,"ptw":30.0,"tbt":52.0,"ind":30.0,"km":4443.0},{"w":33,"d":"23 \u2013 29 Nov 2025","wh":22404,"nm":0.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":66.0,"insp":38.0,"ptw":30.0,"tbt":53.0,"ind":87.0,"km":4483.0},{"w":34,"d":"30 \u2013 06 Des 2025","wh":24916,"nm":0.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":78.0,"insp":134.0,"ptw":30.0,"tbt":52.0,"ind":42.0,"km":4463.0},{"w":35,"d":"07 \u2013 13 Des 2025","wh":27224,"nm":0.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":114.0,"insp":13.0,"ptw":30.0,"tbt":53.0,"ind":29.0,"km":4400.0},{"w":36,"d":"14 \u2013 20 Des 2025","wh":25642,"nm":0.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":86.0,"insp":4.0,"ptw":30.0,"tbt":50.0,"ind":16.0,"km":4321.0},{"w":37,"d":"21 \u2013 27 Des 2025","wh":21888,"nm":0.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":27.0,"insp":8.0,"ptw":30.0,"tbt":51.0,"ind":9.0,"km":4300.0},{"w":38,"d":"28 Des 2025\u2013 03 Jan 2026","wh":21546,"nm":0.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":62.0,"insp":2.0,"ptw":30.0,"tbt":41.0,"ind":6.0,"km":4300.0},{"w":47,"d":"01 Mar 2026 \u2013 07 Mar 2026","wh":35996,"nm":1.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":101.0,"insp":62.0,"ptw":30.0,"tbt":67.0,"ind":19.0,"km":4579.0},{"w":48,"d":"08 Mar 2026 - 14 Mar 2026","wh":34362,"nm":2.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":143.0,"insp":36.0,"ptw":30.0,"tbt":67.0,"ind":7.0,"km":4579.0},{"w":49,"d":"15 Mar 2026 - 21 Mar 2026","wh":25000,"nm":1.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":38.0,"insp":2.0,"ptw":30.0,"tbt":36.0,"ind":3.0,"km":4459.0},{"w":50,"d":"22 Mar 2026 - 28 Mar 2026","wh":21612,"nm":3.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":13.0,"insp":18.0,"ptw":30.0,"tbt":41.0,"ind":23.0,"km":4418.0},{"w":51,"d":"29 Mar 2026 - 04 Apr 2026","wh":44614,"nm":3.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":1.0,"uauc":99.0,"insp":60.0,"ptw":30.0,"tbt":72.0,"ind":84.0,"km":4595.0},{"w":52,"d":"05 Apr 2026 - 11 Apr 2026","wh":51482,"nm":2.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":1.0,"uauc":79.0,"insp":118.0,"ptw":30.0,"tbt":66.0,"ind":66.0,"km":4574.0},{"w":53,"d":"12 Apr 2026 - 18 Apr 2026","wh":58742,"nm":0.0,"fac":0.0,"sec":0,"ed":0.0,"lsr":0.0,"uauc":60.0,"insp":118.0,"ptw":30.0,"tbt":66.0,"ind":66.0,"km":4574.0},{"w":54,"d":"19 Apr 2026 - 25 Apr 2026","wh":60994,"nm":0.0,"fac":0.0,"sec":0.0,"ed":0.0,"lsr":0.0,"uauc":84.0,"insp":3.0,"ptw":30.0,"tbt":78.0,"ind":53.0,"km":4604.0},{"w":55,"d":"26 Apr 2026 - 02 Mei 2026","wh":51850,"nm":0.0,"fac":0.0,"sec":0.0,"ed":0.0,"lsr":0.0,"uauc":42.0,"insp":10.0,"ptw":30.0,"tbt":72.0,"ind":49.0,"km":29412.0},{"w":56,"d":"03 May 2026 - 09 May 2026","wh":67470,"nm":0.0,"fac":0.0,"sec":0.0,"ed":0.0,"lsr":0.0,"uauc":56.0,"insp":11.0,"ptw":65.0,"tbt":77.0,"ind":33.0,"km":2154.0},{"w":57,"d":"10 May 2026 - 16 May 2026","wh":62764,"nm":1.0,"fac":0.0,"sec":0.0,"ed":0.0,"lsr":0.0,"uauc":35.0,"insp":21.0,"ptw":68.0,"tbt":78.0,"ind":68.0,"km":2104.0},{"w":58,"d":"17 May 2026 - 23 May 2026","wh":73102,"nm":0.0,"fac":0.0,"sec":1.0,"ed":0.0,"lsr":0.0,"uauc":38.0,"insp":31.0,"ptw":71.0,"tbt":78.0,"ind":18.0,"km":2532.0},{"w":59,"d":"24 May 2026 - 30 May 2026","wh":52062,"nm":0.0,"fac":0.0,"sec":2.0,"ed":1.0,"lsr":0.0,"uauc":20.0,"insp":14.0,"ptw":76.0,"tbt":78.0,"ind":40.0,"km":1968.0},{"w":60,"d":"31 May 2026 - 06 Jun 2026","wh":67176,"nm":0.0,"fac":0.0,"sec":2.0,"ed":1.0,"lsr":0.0,"uauc":38.0,"insp":31.0,"ptw":77.0,"tbt":78.0,"ind":23.0,"km":2088.0},{"w":61,"d":"07 June 2026 - 13 Jun 2026","wh":68714,"nm":1.0,"fac":0.0,"sec":0.0,"ed":0.0,"lsr":0.0,"uauc":22.0,"insp":44.0,"ptw":78.0,"tbt":77.0,"ind":15.0,"km":2672.0},{"w":62,"d":"14 June 2026 - 20 Jun 2026","wh":69950,"nm":0.0,"fac":1.0,"sec":0.0,"ed":1.0,"lsr":0.0,"uauc":27.0,"insp":44.0,"ptw":70.0,"tbt":108.0,"ind":9.0,"km":2834.0},{"w":63,"d":"21 June 2026 - 27 Jun 2026","wh":65080,"nm":0.0,"fac":1.0,"sec":0.0,"ed":0.0,"lsr":0.0,"uauc":7.0,"insp":13.0,"ptw":60.0,"tbt":107.0,"ind":30.0,"km":2604.0},{"w":64,"d":"28 June 2026 - 04 Jul 2026","wh":66586,"nm":3.0,"fac":0.0,"sec":0.0,"ed":1.0,"lsr":0.0,"uauc":214.0,"insp":16.0,"ptw":57.0,"tbt":103.0,"ind":46.0,"km":2292.0},{"w":65,"d":"05 July 2026 - 11 July 2026","wh":67798,"nm":0.0,"fac":0.0,"sec":0.0,"ed":0.0,"lsr":0.0,"uauc":133.0,"insp":33.0,"ptw":67.0,"tbt":101.0,"ind":11.0,"km":2412.0},{"w":66,"d":"12 July 2026 - 18 July 2026","wh":72504,"nm":1.0,"fac":0.0,"sec":2.0,"ed":0.0,"lsr":0.0,"uauc":184.0,"insp":14.0,"ptw":64.0,"tbt":100.0,"ind":37.0,"km":2964.0},{"w":67,"d":"19 July 2026 - 25 July 2026","wh":64848,"nm":2.0,"fac":0.0,"sec":0.0,"ed":0.0,"lsr":0.0,"uauc":291.0,"insp":6.0,"ptw":35.0,"tbt":99.0,"ind":27.0,"km":2380.0},{"w":68,"d":"26 July 2026 - 01 Aug 2026","wh":72638,"nm":0.0,"fac":0.0,"sec":0.0,"ed":0.0,"lsr":0.0,"uauc":325.0,"insp":20.0,"ptw":85.0,"tbt":102.0,"ind":33.0,"km":2824.0}];
const VENDOR_FULL_NAMES = {
  "EcoOils": "PT EcoOils Jaya Indonesia",
  "EO": "PT EcoOils Jaya Indonesia",
  "GSB": "PT Gerbang Saranabaja",
  "KKS": "PT Kokoh Semesta",
  "PBS": "PT Paramita Bangun Sarana",
  "TSE": "Tera Samudra Engineering",
  "Driver": "Athaya Rent Car Indonesia (ARCI)",
  "PTTS": "PT Prima Tekindo Tirta Sejahtera",
  "YKG": "Yokogawa",
  "MPE": "Mitra Prima Enviro",
  "PME": "Plus Point Engineering Sdn. Bhd.",
  "JGC": "Japan Gasoline Co.",
  "WEN": "PT WIJAYA ENGINDO NUSA",
  "KMS": "PT Karyaindo Makmur Sejati",
  "SEFCON": "SEFCON Engineering Sdn Bhd",
  "SOS": "PT Shield On Service - ALSOK Indonesia",
  "Survey": "CV Karya Jaya Survey (KJS)",
};
const vendorLabel = (code) => {
  const key = (code || "").trim();
  return VENDOR_FULL_NAMES[key] ? `${key} - ${VENDOR_FULL_NAMES[key]}` : key;
};
const SEED_INCIDENTS = [{"id":"EOSM-001-KMS-NM","date":"2025-04-29","time":"10:18","vendor":"KMS","type":"Near Miss","title":"Pile Roll-Off from Trailer During Offloading","summary":"During the offloading of Pile #08 from the trailer, Pile #09, which was positioned on stack #02, unexpectedly rolled towards the trailer side stopper. \nThe force of the rolling pile caused the stopper to break, and Pile #09 subsequently rolled off the stack and onto the ground.\n\nNo injuries were reported as a result of this incident.\n\nImmediate Actions Taken:\n\u2022\tAll offloading activities were immediately stopped.\n\u2022\tAll Permits to Work (PTWs) related to offloading activities were cancelled.","ramActual":"P0, A1, C0, E0\n\nA1: Broken side stopper.","ramPotential":"P4B, A3B, C0B, E0B\n\nP4: Potential to cause permanent disability or fatalities if the rolling pile had struck someone.\n\nA3: A more severe outcome could have led to moderate asset damage and causing operational delays.","immediateCause":"Inadequate Chocking: Failed to prevent unexpected pile movement. \n\nInsufficiently Robust Stopper: Unable to withstand rolling pile force.\n\nImproper Offloading Sequence: Created instability allowing the roll.","rootCause":"Lack of clear or effectively implemented safe work procedures for chocking and offloading, coupled with potential gaps in inspection and supervision, resulted in the event for the pile to roll and fall.\n \nLess than adequate risk assessment failed to identify and mitigate these hazards thus resulting the roll off of the pile.\n\nInadequate design of the pile stoppers, which underestimated the potential forces generated during loading and unloading operations. This underestimation failed to account for the combined effects of the pile's weight and its momentum, resulting in the stoppers being incapable of preventing the pile from falling.","lessonsLearned":"Robust Chocking and Offloading Sequence: Inadequate chocking and an improper offloading sequence directly led to the pile rolling and falling from the trailer.\n\nThorough Pre-Task Risk Assessments: Less than adequate pre-task risk assessments failed to identify and address underlying issues like insufficient chocking and improper offloading sequences, which contributed to the incident."},{"id":"EOSM-002-PBS-NM","date":"2025-06-03","time":"15:30","vendor":"PBS","type":"Near Miss","title":"Unauthorized Tandem Lift of Rebars","summary":"Unauthorized tandem offloading of rebar stacks/bundles using \na FOCO crane and excavator was performed without a JHA, PTW, or lifting plan. \nThis activity was not authorized and lacked the required PTW for tandem lifts. \n\nNote: While a PTW covered FOCO crane-only offloading, tandem lifts were not assessed or permitted.\n\nNo injuries occurred during the activity. \n            \nImmediate Actions Taken:\n\u2022\tLifted load immediately lowered/rested to the ground from its mid-air position.\n\u2022\tStop Work Obligation exercised.\n\u2022\tActivity suspended pending investigation.","ramActual":"P0, A0, C0, E0\n\nNo injuries occurred during the activity.","ramPotential":"P4B, A3B, C0B, E0B\n\nP4B: A failed tandem lift carries the potential for permanent disability or fatal injuries.\n\nA3B: A worse scenario could have resulted in moderate asset damage and operational delays.","immediateCause":"Unsafe Act:\nOperation of equipment (FOCO crane and excavator) without proper authorization for a tandem lift and using the excavator improperly for this unintended purpose.\nThe lifting itself was performed improperly, not optimally according to the equipment's function.\n\nUnsafe Condition: \nThe activity was conducted without the required safety documentation, specifically a JHA & PTW for tandem lifts, or a lifting plan.\nThere was a lack of adequate or appropriate equipment for the task, leading to the improper use of the excavator.","rootCause":"Management failed to ensure effective oversight, allowing an unauthorized and un-planned high-risk activity to be initiated.\n\nManagement did not ensure the availability of adequate and appropriate tools and equipment for the task, leading to the improper use of existing machinery.\n\nManagement's system lacked comprehensive or enforced work standards, as evidenced by the absence of critical safety documentation (JHA, PTW, lifting plan) for the complex tandem lift.","lessonsLearned":"Management must ensure the integrity and enforcement of the safety management system by proactively establishing and maintaining robust work standards, procedures (including for non-standard operations like tandem lifts), and by allocating appropriate resources and equipment to prevent unauthorized deviations.\n\nDirect supervision must be empowered and held accountable for strictly enforcing safety protocols, ensuring all high-risk activities are properly authorized and planned before execution, and intervening immediately to stop unsafe acts.\n\nFor all complex or non-routine tasks, a comprehensive JHA, PTW, and specific lifting plans are mandatory. All equipment must be thoroughly assessed and confirmed suitable for its intended (and any non-standard) use to prevent improvisation or misuse."},{"id":"EOSM-003-PBS-NM","date":"2025-06-13","time":"15:30","vendor":"PBS","type":"Near Miss","title":"Subcontractor PPE Violation","summary":"EcoOils Site Lead observed a trailer driver, whose vehicle had been parked in the laydown area since 09:00Hrs awaiting unloading, performing repairs on his trailer without the required PPE. specifically, the driver was not wearing a hard hat, safety glasses, or safety shoes, and was wearing shorts with his shirt removed.","ramActual":"P0, A0, C0, E0\nNo injuries occurred during the activity.","ramPotential":"P2B, A0, C0, E0\nModerate to severe head, eye, or crushing injuries given the complete lack of essential PPE during trailer repairs.","immediateCause":"Unsafe Acts:\nThe driver failed to use personal protective equipment as required for work.","rootCause":"Management had not sufficiently prioritized or allocated resources for comprehensive, ongoing safety training for security staff, specifically concerning their role in PPE compliance and pre-entry safety socialization. This led to security officers lacking the training and dedicated time to adequately socialize PPE use, which in turn resulted in the driver failing to use required PPE.","lessonsLearned":"Proactive Safety Culture: Management must prioritize and adequately resource continuous safety training and clear communication for all personnel, particularly those responsible for enforcing safety standards.\n\nImplement Active Compliance Controls: Move beyond passive checks to implement robust controls, including detailed checklists, thorough training, real-time communication, and regular audits to ensure consistent safety adherence."},{"id":"EOSM-004-PBS-NM","date":"2025-06-27","time":"14:00","vendor":"PBS","type":"Near Miss","title":"Use of Vibratory Roller Without Critical Safety Controls & Risk Mitigation Measures","summary":"Incident Summary:\nThe DDR60CS Vibratory Roller, which has been identified as having an exposed rotating wheel, was used for compaction work without adhering to Critical Safety Controls & Risk Mitigation Measures\n\nThe following temporary risk mitigation measures (below trail email highlighted in yellow), as requested for the use of the Vibratory Roller, were not implemented:\n\u2022\tThe task was authorized by the supervisor without a PTW and JHA.\n\u2022\tThe work area was not cordoned off with a safety barrier.\n\u2022\tFour individuals were present within the unguarded work zone, in close proximity to the operating equipment.\n\u2022\tThe roller was operated by personnel who had not been identified as authorized users.\n\u2022\tThe crew was not briefed on the specific hazards of the rotating equipment.\n\u2022\tThe equipment lacked a valid inspection tag.\n\nNo injuries occurred during the activity. \n            \nImmediate Actions Taken:\n\u2022\tAll operations involving the DDR60CS Vibratory Roller were stopped.\n\u2022\tA \"Do Not Operate\" tag was placed on the DDR60CS Vibratory Roller.","ramActual":"P0, A0, C0, E0\n\nNo injuries, asset damage, environmental effects, or reputational impact.","ramPotential":"P4B, A0B, C0B, E0B\n\nP4B: Potential for Permanent Total Disability. This potential incident has been heard of in the industry","immediateCause":"Unsafe Acts:\nOperating Equipment Without Authority\nUsing Defective Equipment\nFailure to Follow Procedure\n\nUnsafe Conditions:\nExposed rotating part (un-guarded rotating wheel) on the DDR60CS Vibratory Roller.\nUnguarded work zone / work area not cordoned off.\nEquipment lacked a valid inspection tag.","rootCause":"Ineffective management system for cascading, understanding, and enforcing safety directives, especially for critical hazards.\n\nGaps in safety culture, leading to inconsistent prioritization or verification of procedure adherence by higher management.\n\nInsufficient management oversight or auditing of supervisor adherence to safety protocols and risk control implementation.\n\nLack of a robust management verification process to confirm safety improvements for high-risk equipment are in place before use.","lessonsLearned":"Non-Negotiable Adherence to PTW and Safety Controls: Bypassing Permit to Work and Job Hazard Analysis protocols, and failing to implement critical safety controls, directly led to the near miss and exposed personnel to a high risk of severe injury.\n\nAccountability and Verification in Safety Management: Ineffective management oversight and a lack of robust verification processes allowed unsafe equipment use despite known hazards and communicated mitigation measures, highlighting a critical breakdown in safety culture and enforcement."},{"id":"EOSM-005-KMS-PD","date":"2025-07-01","time":"14:20","vendor":"KMS","type":"Property Damage","title":"Pile Roll-Off from Trailer During Offloading and Damaged","summary":"Incident Summary:\nWhile the 14-ton TMC was engaged in offloading piles from the trailer. \nAs the second pile was being lifted from Stack #03, the fourth pile from the same stack rolled off towards the side of the trailer. \nThe fallen pile was damaged and cracked upon impact.\n\nNote:- No injuries occurred during the activity.\n           \nImmediate Actions Taken:\n\u2022\tThe activity was temporarily stopped.","ramActual":"P0, A1C, C0, E0\n\nA1C:- The fallen pile was damaged and cracked upon impact. (Cost of Pile approx 5 Million IDR)","ramPotential":"P4B, A3B, C0B, E0B\n\nP4B: Potential to cause permanent disability or fatalities if the rolling pile had struck someone.\n\nA1C: Same as actual","immediateCause":"","rootCause":"Insufficient space and simultaneous operations: The primary cause was the narrow work area. The presence of two cranes working at the same time and in a confined space limited the TMC's ability to open its boom fully.\n\nInadequate pile security: The wooden beams used to secure the piles were too small and not strong enough to withstand the force of a nearby pile being lifted. The existing safety awning and side iron were also not robust enough.\n\nImproper lifting sequence: The lifting process caused a shifting force on the remaining piles on the trailer, leading to one falling.","lessonsLearned":"\u2022\tImportance of Proper Workspace and Equipment\nA major contributing factor to the incident was the narrow work area which restricted the crane's operation. This highlights the need for a thorough site and equipment assessment to ensure all machinery can operate with adequate clearance and a full range of motion. The report also notes the inadequacy of the wooden beams used for securing the piles. This shows the importance of using appropriately sized and robust equipment for the task, especially for safety-critical components.\n\n\u2022\tNeed for Robust Safety Protocols and Oversight\nThe incident was a direct result of a failure in a key safety protocol. The lifting process itself, combined with the inadequate securing of the piles, created a hazard. This points to a need for stricter adherence to and continuous monitoring of safety protocols, such as the Job Hazard Analysis (JHA) and Job Method Statement (JMS). The report's corrective actions, like having continuous inspections and a new lifting sequence, reinforce the importance of proactive and ongoing oversight to prevent similar incidents."},{"id":"EOSM-006-PBS-FAC","date":"2025-07-14","time":"1430 Hrs.","vendor":"PBS","type":"First Aid Case","title":"Minor Scratch on Right Shin","summary":"\uf02d\tIP sustained a scratch to their right shin after inadvertently stepping on a spun pile hole. The plywood/board cover was inadequately positioned and partially concealed by a sand pile, failing to prevent inadvertent entry.\n\n\uf02d\tOn-site first aid was administered to the IP.\n\uf02d\tIP was sent to a clinic for medical assessment, including a tetanus injection.\n\uf02d\tSite survey was conducted to identify and secure other potential pile openings.\n\uf02d\tSafety stand-down was held with site workers, supervisors, and HSSE (15th July 2025).","ramActual":"P1B, A0, C0, E0\n\nP1B \u2013 Minor scratch on right shin.","ramPotential":"P2B, A0, C0, E0\n\nP2B: - Personnel could potentially sustain a minor injury categorized as an MTC, or a LWC or RWC with a duration of up to five days.","immediateCause":"","rootCause":"Ineffective safety communication.\n\nPoor enforcement of hazard protocols.\n\nInadequate supervisory verification.\n\nConvenience overrode formal procedures.\n\nWeak system to ensure consistent compliance.","lessonsLearned":"The importance of effective communication, consistent enforcement, and active supervisory verification of safety directives and hazard controls. Despite general safety measures, inadequate communication, an inadequate safety culture (where convenience overrode procedures), and a lack of supervisory verification directly led to an unaddressed hazard and subsequent injury.\n\nThe necessity for robust operational leadership and systemic enforcement mechanisms in safety management. The investigation points to inadequate operational leadership in monitoring hazards and addressing at-risk behaviours, alongside a weak system for enforcing safety procedures, as fundamental causes. This emphasizes that strong leadership and a resilient system are vital to ensure consistent compliance and prevent similar occurrences."},{"id":"EOSM-007-PBS-NM","date":"2025-08-18","time":"09:30","vendor":"PBS","type":"Near Miss","title":"Canteen Worker Trip and Fall","summary":"","ramActual":"P0, A0, C0, E0\n\nCanteen worker fall with one foot in hole of the drainage for the carwash of ecooils cars that already cover by plywood","ramPotential":"P3B, A0B, C0B, E0B\n\nP3:  Tripping over a foot can potentially lead to injury","immediateCause":"","rootCause":"Insufficient risk assessment and control measures: The investigation found that the hole for the drainage was covered with plywood, which was an inadequate control measure for the hazard. A more thorough risk assessment should have identified the need for a more secure and permanent cover to prevent falls.\n\nUnfocused behavior of the canteen worker: The \"Why Causal Reasoning\" section notes that the canteen worker was \"not foccused when walk in that area,\" which contributed to them not noticing the inadequate cover and stepping on it.","lessonsLearned":"Plywood is not a durable or reliable cover for a manhole. While it may provide a temporary solution, it poses a significant tripping hazard and can easily break or shift, leading to a fall.\n\nA lack of situational awareness can lead to injury. The canteen worker was reportedly \"not focused\" when walking in the area, which contributed to the incident.\n\nProactive risk assessment and control measures are critical. The underlying cause was identified as insufficient risk assessment and control measures. Had a proper assessment been performed, a more permanent and safer solution for the manhole cover would have been implemented. This would have prevented the incident from happening altogether.\n\nThis incident highlights the importance of providing durable and secure covers for holes, emphasizing that temporary fixes are often insufficient and can create new hazards. It also serves as a reminder for all personnel to maintain situational awareness in their work environment to prevent similar incidents."},{"id":"EOSM-008-PBS-NM","date":"2025-09-02","time":"10:00","vendor":"PBS","type":"Near Miss","title":"Unsafe Lifting Operation and Working at Height","summary":"Incident Summary:\nWhile offloading a stack of scaffold boards with a 14-ton FOCO crane, a rigger\u2019s helper climbed onto the partially elevated load without a safety harness and attempted to attach a second rigging sling.\n \nLSR Violation:\n\u2022\tWorking at Height\n\u2022\tLine of Fire            \n \nImmediate Actions Taken:\n\u2022\tStop Work Authority was exercised.\n\u2022\tThe Permit to Work (PTW) was withdrawn and the activity was suspended.\n\u2022\tA safety stand-down was conducted with the involved personnel.","ramActual":"P0B, A0B, C0A, E0B","ramPotential":"P3B, A0B, C0B, E0B","immediateCause":"","rootCause":"Insufficient risk assessment and control measures: The investigation determined that there was a lack of proper planning and a failure to identify and mitigate the risks associated with the lifting operation, leading to a dangerous work procedure.\n\nLack of training and understanding: The workers involved did not have a full understanding of safe lifting procedures, as evidenced by one worker climbing on the material being lifted to \"secure\" it. There was a clear lack of training or communication regarding the correct way to perform the task.\n\nAbsence of key safety documentation: The JSA and MS were not on-site, indicating a breakdown in the process for ensuring that all personnel understood the safety protocols for the job before work began.","lessonsLearned":"Never climb on lifted materials without proper safety measures. A worker climbed on the scaffolding material being lifted without a harness, which is an extremely dangerous and unsafe act that could have resulted in a fatal fall.\n\nAlways ensure that all safety documents, such as the Job Safety Analysis (JSA) and Method Statement (MS), are on-site and understood by all involved personnel before starting a job. The work was stopped because these critical documents were not present, indicating a failure in pre-work planning and communication"},{"id":"EOSM-009-EO-NM","date":"2025-09-17","time":"18:30 Hrs","vendor":"EO","type":"Near Miss","title":"Electrical Shock from Office Container Door Handle","summary":"Incident Summary:\nAn EcoOils representative was on site to supervise an extended shift concrete pouring activity by a contractor.\n\nAlthough rain had commenced around 5 PM, at approximately 6:30 PM, the representative felt an electric shock while attempting to open the container office door and immediately released the handle.\n\nThe individual was unharmed and did not require medical attention.\n\nAction Taken\n\u2022\tAn investigation was initiated this morning by site management with support from an E&I representative.\n\u2022\tThe PBS contractor has been requested to check the electrical installation in the EcoOils office.","ramActual":"P0B, A0B, C0A, E0B","ramPotential":"P3B, A0B, C0B, E0B","immediateCause":"","rootCause":"Lightning Surcharge: The incident occurred during rain, suggesting that lightning could have caused an electrical surge or indirect strike that led to the electric shock. This is a common phenomenon in electrical systems exposed to storms.\n\nInadequate Grounding: The electric shock a person felt upon touching a metal object (the container office door) strongly indicates that the electrical system was not properly grounded. Proper grounding would have safely dissipated the electrical charge into the earth, preventing it from accumulating on the container's metal frame.","lessonsLearned":"Temporary electrical setups are a significant hazard, especially in wet conditions. The incident shows that the container office's electrical system was not properly isolated or grounded, and the rain created a pathway for the current to flow, causing an electric shock. This highlights the dangers of using temporary or uninspected electrical systems.\n\nProactive checks and immediate action can prevent serious injury or fatality. The individual immediately released the handle and was unharmed. This and the immediate investigation that followed prevented a potentially more severe incident. This demonstrates the importance of a swift response to safety concerns and a culture where personnel are trained to react correctly in such situations"},{"id":"EOSM-010-PBS-NM","date":"2025-09-23","time":"15:00 Hrs","vendor":"PBS","type":"Near Miss","title":"Employee Altercation","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-011-PBS-NM","date":"2025-09-28","time":"14:00","vendor":"PBS","type":"Near Miss","title":"Unauthorized Asphalt Preparation Activity","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-012-PBS-NM","date":"2025-09-30","time":"14:30","vendor":"PBS","type":"Near Miss","title":"Supervisor Smoking in the Warehouse","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-013-PBS-NM","date":"2025-10-01","time":"00:00","vendor":"PBS","type":"Near Miss","title":"Collapse of Zinc & Wooden Canteen Structure Due to Extreme Weather","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-014-PBS-NM","date":"2025-10-01","time":"00:00","vendor":"PBS","type":"Near Miss","title":"Rebar Tower Lean Due to Strong Winds","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-015-KKS-NM","date":"2025-10-07","time":"10:00","vendor":"KKS","type":"Near Miss","title":"Dislodged Scaffold Material on Trailer and Offloading Without Permit","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-016-GSB-NM","date":"2025-10-09","time":"08:30","vendor":"GSB","type":"Near Miss","title":"Webbing Stuck on Lightning Arrestor","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-017-PBS-PD","date":"2025-10-21","time":"14:00","vendor":"PBS","type":"Property Damage","title":"Anchor Bolts Damaged by Excavator","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-018-PBS-NM","date":"2025-10-21","time":"14:00","vendor":"PBS","type":"Near Miss","title":"Rebar Column Tilted Due to Premature Support Removal","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-019-PBS-NM","date":"2025-10-21","time":"16:35","vendor":"PBS","type":"Near Miss","title":"Employee Driving Pick-Up Without Driving License","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-020-KKS-NM","date":"2025-10-22","time":"11:10","vendor":"KKS","type":"Near Miss","title":"PPE Non-Compliance, Smoking in Restricted Area, and Duty to Intervene Failure","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-021-PBS-NM","date":"2025-10-29","time":"08:45","vendor":"PBS","type":"Near Miss","title":"Potential Dropped Object During Lifting","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-022-PBS-FAC","date":"2025-10-29","time":"14:30","vendor":"PBS","type":"First Aid Case","title":"Worker Experienced Dizziness","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-023-GSB-NM","date":"2025-10-30","time":"17:00","vendor":"GSB","type":"Near Miss","title":"Dropped Object","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-024-KKS-NM","date":"15th Nov 2","time":"15:00","vendor":"KKS","type":"Near Miss","title":"Unapproved Life-Line Beam Attachments and Work at Height Violations","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-025-GSB-NM","date":"2025-11-17","time":"15:00","vendor":"GSB","type":"Near Miss","title":"Crane Left Running & Unsupervised During Break","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-026-PBS-NM","date":"2025-11-18","time":"10:09","vendor":"PBS","type":"Near Miss","title":"Dump Truck Bucket Contact with Unpowered Overhead Cable","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-027-KKS-NM","date":"2026-01-12","time":"17:00","vendor":"KKS","type":"Near Miss","title":"Working at Height Violation","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-028-PBS-NM","date":"2026-01-14","time":"09:48","vendor":"PBS","type":"Near Miss","title":"Excavator Contacted Temporary Roof","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-029-KKS-SEC","date":"2026-01-26","time":"15:24","vendor":"KKS","type":"Security","title":"Unauthorized Vehicle Entry to Site","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-030-KKS-NM","date":"2026-01-26","time":"15:24","vendor":"KKS","type":"Near Miss","title":"Failure to Maintain Safe Lifting Radius","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-031-WEN-SEC","date":"2026-01-28","time":"10:00","vendor":"WEN","type":"Security","title":"Suspected Unauthorized Removal of Grounding Cable","summary":"E#01 \u2013 Silo & EM Shed Area | 28 January 2026, 10:00 Hrs.\nSite personnel discovered the unauthorized removal of 32 meters of 70sqmm insulated copper grounding cable at the Silo and EM Shed area. The cables, installed on 26 January 2026, were cut and extracted from the ground by unknown parties. Irregular excavation marks were observed at the site, indicating deliberate, targeted extraction.\n\nE#02 \u2013 Hexane Bund Wall / Tank Farm | 04 February 2026, 14:30 Hrs.\nPT. WEN supervisor and team discovered during an afternoon inspection that newly installed grounding cable had been stolen from inside the hexane bund wall area. A subsequent sweep of the Tank Farm confirmed a total of 26 meters of cable had been cut and removed.\n\nE#03 \u2013 Pipe Sleeper #00, Extraction Plant to Tank Farm | 23 February 2026, 13:45 Hrs.\nDuring a scheduled inspection, site personnel identified the theft of 12 meters of grounding wire from the Pipe Sleeper #00 section spanning the vicinity between the Extraction Plant and the Tank Farm. This incident confirms a continued pattern of unauthorized cable removal on site.\n\nE#04 \u2013 Pipe Rack 001 | 28 February 2026, 08:30 Hrs.\nAn additional 20 meters of grounding wire was found stolen at Pipe Rack 001. Discovery was made during morning site activities. This represents the fourth confirmed incident of cable theft within a 31-day period.","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-032-KKS-NM","date":"2026-02-11","time":"15:33","vendor":"KKS","type":"Near Miss","title":"Smoldering Plastic Bag at Pipe Rack #05","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-033-GSB-NM","date":"2026-02-12","time":"07:40","vendor":"GSB","type":"Near Miss","title":"Initiation of Grinding Activities Without Valid PTW Authorization","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-034-PBS-NM","date":"2026-02-13","time":"14:26","vendor":"PBS","type":"Near Miss","title":"Operating Vehicle While Using a Mobile Phone Without Flagman","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-035-KKS-NM","date":"2026-02-13","time":"16:20","vendor":"KKS","type":"Near Miss","title":"PPE Non-Compliance and Fall Protection Breach","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-036-KKS-NM","date":"2026-02-23","time":"17:00","vendor":"KKS","type":"Near Miss","title":"Trailer Contact with Overhead Unpowered Electrical Cable","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-037-WEN-NM","date":"2026-02-22","time":"15:00","vendor":"WEN","type":"Near Miss","title":"PTW Violation in the Tank Farm","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-038-KKS-NM","date":"2026-02-26","time":"15:34","vendor":"KKS","type":"Near Miss","title":"Unauthorized Crane Mobilization and Lifting Operation","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-039-KKS-NM","date":"2026-03-02","time":"16:28","vendor":"KKS","type":"Near Miss","title":"Oxygen Hose Failure During Cylinder Changeover","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-040-PBS-NM","date":"2026-03-09","time":"","vendor":"PBS","type":"Near Miss","title":"Unauthorized   Energization of Sub Distribution Board Without Permit and LOTO","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-041-KKS-NM","date":"2026-03-09","time":"","vendor":"KKS","type":"Near Miss","title":"Dropped Scaffolding Clamp","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-42-KKS-NM","date":"2026-03-17","time":"","vendor":"KKS","type":"Near Miss","title":"Equipment contact anchor bolt during installation","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-43-WEN-NM","date":"2026-03-23","time":"","vendor":"WEN","type":"Near Miss","title":"Gratting removal and modification without PTW","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-44-PBS-NM","date":"2026-03-26","time":"","vendor":"PBS","type":"Near Miss","title":"Crew Working Under Suspended Load","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-45-KKS-NM","date":"2026-03-26","time":"","vendor":"KKS","type":"Near Miss","title":"Dropped Tool   from Scaffold \u2014 Wire Cutter Fall from Height","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-046-KKS-NM","date":"2026-03-30","time":"","vendor":"KKS","type":"Near Miss","title":"Eye Protection   Non-Compliance During Magnetic Drill Operation","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-47-GSB-NM","date":"2026-03-30","time":"","vendor":"GSB","type":"Near Miss","title":"Concurrent Spray   Painting and Hot Work at Tank and Pipe Rack","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-48-PBS-NM","date":"2026-04-04","time":"","vendor":"PBS","type":"LSR","title":"PTW LSR Breach \u2014 Stop Work Order Countermanded","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-49-KKS-NM","date":"2026-04-08","time":"","vendor":"KKS","type":"LSR","title":"LSR Violation Working at Height and Working without PTW","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-50-PBS-NM","date":"2026-04-10","time":"","vendor":"PBS","type":"Near Miss","title":"Truck Loaded with Concrete Blocks Stuck in Mud \u2013 Unsafe Recovery Attempt Without Re-Assessment","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-51-KKS-NM","date":"2026-05-15","time":"","vendor":"KKS","type":"Near Miss","title":"Crane struck scaffolding bridge at Pipe Rack Bridge","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-52-WEN-SEC","date":"22/23-May-","time":"","vendor":"WEN","type":"Security","title":"Un-authorized Removal of Field Earthing Busbar Copper","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-53-WEN-PD","date":"2026-05-28","time":"","vendor":"WEN","type":"Property Damage","title":"Ex Power Outlet Damaged","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-54-GSB-SEC","date":"2026-05-29","time":"","vendor":"GSB","type":"Security","title":"Theft of Electric Welding Cable","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-55-WEN-SEC","date":"2026-05-30","time":"","vendor":"WEN","type":"Security","title":"Un-authorized Removal of Field Earthing cable & Copper Rod","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-56-PBS-NM","date":"2026-06-12","time":"","vendor":"PBS","type":"Near Miss","title":"Smoldering on formworks at SBE Warehouse","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-57-KKS-FAC","date":"2026-06-19","time":"","vendor":"KKS","type":"First Aid Case","title":"Electric Shock   to worker during Portable Electrode Dryer Connection","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-58-PBS-PD","date":"2026-06-15","time":"","vendor":"PBS","type":"Property Damage","title":"Crane Front   Guard Contact with Fire Water Line during Maneuvering","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-59-KKS-FAC","date":"2026-06-27","time":"","vendor":"KKS","type":"First Aid Case","title":"Scaffolder experienced dizziness","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-60-EO-NM","date":"2026-06-29","time":"","vendor":"EO","type":"Near Miss","title":"Dislodged container false ceiling","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-61-KKS-NM","date":"2026-07-02","time":"11:28","vendor":"KKS","type":"Near Miss","title":"PTW Scope Deviation \u2013 Unauthorized Roofing Activity above Elevation 17K","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-62-EO-PD","date":"2026-07-04","time":"05:30","vendor":"EO","type":"Property Damage","title":"CCTV Cable Damaged by Delivery Truck at West Gate","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-63-KKS-NM","date":"2026-07-04","time":"","vendor":"KKS","type":"Near Miss","title":"Dropped Object - Scaffolding Tube","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-64-PBS-SEC","date":"2026-07-13","time":"17:30","vendor":"PBS","type":"Security","title":"Copper Cable Bundles taken out from site without authorization","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-65-GSB-SEC","date":"2026-07-14","time":"11:30","vendor":"GSB","type":"Security","title":"Welding Cable Copper Core theft","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-66-PBS-NM","date":"2026-07-17","time":"07:25","vendor":"PBS","type":"Near Miss","title":"Employee on site without induction","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-67-GSB-NM","date":"2026-07-20","time":"02:00","vendor":"GSB","type":"Near Miss","title":"Deviation from Approved Method Statement and Lack of Readiness Planning During Upending of Silo Cone","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-68-WEN-PD","date":"2026-07-22","time":"16:00","vendor":"WEN","type":"Property Damage","title":"Underground Cable Damage During Excavation Activity","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-69-WEN-PD","date":"2026-07-29","time":"10:00","vendor":"WEN","type":"Property Damage","title":"Temperature Gauges Damaged","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-70-KKS-NM","date":"2026-08-20","time":"10:12","vendor":"KKS","type":"Near Miss","title":"Improper Use of Wire Rope to Secure Pipes for Fit-up","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""},{"id":"EOSM-71-KKS-NM","date":"2026-08-27","time":"10:18","vendor":"KKS","type":"Near Miss","title":"Unauthorized Telehandler Maintenance, Site Access Breach, & Oil Spill","summary":"","ramActual":"","ramPotential":"","immediateCause":"","rootCause":"","lessonsLearned":""}];

const SEED_VERSION = "2026-09-01-digital-font";
const LOGO_DATA_URI = "data:image/png;base64,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";
const LEADING_LAGGING_IMG = "data:image/png;base64,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";
const PROJECTS = [{ id: "mentari", name: "Mentari Project", locations: ["Sei Mangkei SEZ (Site)", "Offsite / Office"] }];
const INCIDENT_TYPE_COLORS = { "Near Miss": "#FFB020", "Property Damage": "#8E7CC3", "Security": "#E5533D", "First Aid Case": "#3FB8AF", "LSR": "#D6336C" };



const LAGGING_METRICS = [
  { key: "trcf", label: "Total Recordable Case Frequency (TRCF)", unit: "" },
  { key: "ltif", label: "Total Lost Time Injury Frequency (LTIF)", unit: "" },
  { key: "lsr", label: "LSR Incidents", unit: "" },
  { key: "otherSig", label: "Other Significant Incidents (RAM4+)", unit: "" },
  { key: "fac", label: "First Aid Cases (FAC)", unit: "" },
  { key: "oi", label: "Occupational Illness (OI)", unit: "" },
  { key: "nearMiss", label: "Near Miss Incidents (NM)", unit: "" },
  { key: "envIncident", label: "Environmental Incidents (EI)", unit: "" },
  { key: "equipDamage", label: "Equipment / Property Damage", unit: "" },
  { key: "security", label: "Security Incidents", unit: "" },
  { key: "mvi", label: "Motor Vehicle Incidents (Site)", unit: "" },
  { key: "transport", label: "Transport Incidents (Road/Air/Water/Rail)", unit: "" },
];

const LEADING_METRICS = [
  { key: "uaUc", label: "UA / UC Reported", unit: "" },
  { key: "mgmtVisits", label: "Management Visits", unit: "" },
  { key: "walkthroughs", label: "Walkthroughs", unit: "" },
  { key: "inspections", label: "Inspections / Surveys", unit: "" },
  { key: "ptw", label: "PTW Issued", unit: "" },
  { key: "jha", label: "JHAs Developed/Reviewed", unit: "" },
  { key: "liftPlans", label: "Lift Plans", unit: "" },
  { key: "inductions", label: "Inductions", unit: "" },
  { key: "trainings", label: "Trainings Conducted", unit: "" },
  { key: "toolboxTalks", label: "Toolbox Talks", unit: "" },
  { key: "meetings", label: "Meetings", unit: "" },
  { key: "drills", label: "Drills", unit: "" },
  { key: "audits", label: "Audits", unit: "" },
  { key: "moc", label: "MOCs Raised", unit: "" },
  { key: "standDowns", label: "Stand Downs", unit: "" },
  { key: "drugAlcohol", label: "Drug & Alcohol Tests", unit: "" },
  { key: "km", label: "Kilometers Driven", unit: "km" },
  { key: "warnings", label: "Site HSSE Violation Warnings", unit: "" },
];

const WASTE_METRICS = [
  { key: "hazWaste", label: "Hazardous Waste", unit: "t" },
  { key: "nonHazWaste", label: "Non-Hazardous Waste", unit: "t" },
];

/* ---------------------------------------------------------------------- */
/* Helpers                                                                 */
/* ---------------------------------------------------------------------- */

const sum = (obj) => Object.values(obj).reduce((a, b) => a + (Number(b) || 0), 0);
const fmt = (n) => Math.round(n).toLocaleString("en-US");
const fmtDecimal2 = (n) => (n || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmt1 = (n) => (Math.round(n * 10) / 10).toLocaleString("en-US");
const fmt2 = (n) => (Math.round((n || 0) * 100) / 100).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const monthLabel = (m) => {
  const [y, mo] = m.split("-");
  return new Date(Number(y), Number(mo) - 1, 1).toLocaleString("en-US", { month: "short", year: "2-digit" });
};
const emptyCats = (cats) => Object.fromEntries(cats.map((c) => [c, 0]));

const PALETTE = {
  bg: "#0E1418",
  panel: "#161D23",
  panelAlt: "#1C242B",
  border: "#2A343C",
  amber: "#FFB020",
  amberDim: "#7A5A1E",
  teal: "#3FB8AF",
  red: "#E5533D",
  purple: "#8E7CC3",
  text: "#E9EDEF",
  textDim: "#8B979E",
};

const SITE_COLOR = "#FFB020";
const OFFSITE_COLOR = "#3FB8AF";
const CONTRACTOR_COLORS = ["#FFB020","#3FB8AF","#E5533D","#8E7CC3","#6FA8DC","#E06C75","#82C99A","#D4A574","#5C9EAD","#C97B84","#A3BE8C","#B08968"];
const ADMIN_USERNAME = "iggifernandes@gmail.com";
const ADMIN_PASSWORD = "0822";

export default function App() {
  const [months, setMonths] = useState(SEED_MONTHS);
  const [incidents, setIncidents] = useState(SEED_INCIDENTS);
  const [customVendors, setCustomVendors] = useState({ site: [], offsite: [] });
  const [loaded, setLoaded] = useState(false);
  const [showEntry, setShowEntry] = useState(false);
  const [showWorkHoursDetail, setShowWorkHoursDetail] = useState(false);
  const [showScorecardDetail, setShowScorecardDetail] = useState(false);
  const [project, setProject] = useState(PROJECTS[0].id);
  const [location, setLocation] = useState(PROJECTS[0].locations[0]);
  const [toast, setToast] = useState(null);
  const [tick, setTick] = useState(0);
  const fileInputRef = useRef(null);
  const [adminUnlocked, setAdminUnlocked] = useState(false);
  const [showAccessGate, setShowAccessGate] = useState(false);
  const [pendingAction, setPendingAction] = useState(null);
  const [accessLog, setAccessLog] = useState([]);
  const [showAccessLog, setShowAccessLog] = useState(false);

  useEffect(() => {
    (async () => {
      let storedMonths = null;
      let storedIncidents = null;
      let storedCustomVendors = null;
      let storedVersion = null;
      try {
        const res = await window.storage.get("workhours-months", false);
        if (res && res.value) storedMonths = JSON.parse(res.value);
      } catch (e) {}
      try {
        const res4 = await window.storage.get("workhours-incidents", false);
        if (res4 && res4.value) storedIncidents = JSON.parse(res4.value);
      } catch (e) {}
      try {
        const res5 = await window.storage.get("workhours-custom-vendors", false);
        if (res5 && res5.value) storedCustomVendors = JSON.parse(res5.value);
      } catch (e) {}
      try {
        const res3 = await window.storage.get("hsse-seed-version", false);
        if (res3 && res3.value) storedVersion = res3.value;
      } catch (e) {}

      const normalizeMonth = (m) => ({ ...m, staffPermanent: m.staffPermanent || { site: 0, offsite: 0 }, indicators: m.indicators || {} });

      if (storedCustomVendors && typeof storedCustomVendors === "object") {
        setCustomVendors({ site: storedCustomVendors.site || [], offsite: storedCustomVendors.offsite || [] });
      }

      if (storedVersion === SEED_VERSION) {
        // Up to date — use whatever is stored (includes any manual edits).
        if (Array.isArray(storedMonths) && storedMonths.length) setMonths(storedMonths.map(normalizeMonth));
        if (Array.isArray(storedIncidents)) setIncidents(storedIncidents);
      } else {
        // Seed data has been refreshed since this browser last loaded —
        // bring in the new import, but keep any manually-added records.
        const manualIncidents = Array.isArray(storedIncidents) ? storedIncidents.filter((i) => !SEED_INCIDENTS.some((s) => s.id === i.id)) : [];
        const mergedIncidents = [...SEED_INCIDENTS, ...manualIncidents];

        const REMOVED_MONTHS = ["2026-08"];
        const monthMap = new Map(SEED_MONTHS.map((m) => [m.month, normalizeMonth(m)]));
        if (Array.isArray(storedMonths)) {
          storedMonths.forEach((m) => {
            if (REMOVED_MONTHS.includes(m.month)) return;
            if (monthMap.has(m.month)) {
              // Keep the fresh seed's vendor data, but preserve any staffPermanent/indicators the user already entered.
              const existing = monthMap.get(m.month);
              const preserved = { ...existing };
              if (m.staffPermanent && (m.staffPermanent.site || m.staffPermanent.offsite)) preserved.staffPermanent = m.staffPermanent;
              if (m.indicators && Object.keys(m.indicators).length) preserved.indicators = m.indicators;
              monthMap.set(m.month, preserved);
            } else {
              monthMap.set(m.month, normalizeMonth(m));
            }
          });
        }
        const mergedMonths = Array.from(monthMap.values());

        setMonths(mergedMonths);
        setIncidents(mergedIncidents);
        try { await window.storage.set("workhours-months", JSON.stringify(mergedMonths), false); } catch (e) {}
        try { await window.storage.set("workhours-incidents", JSON.stringify(mergedIncidents), false); } catch (e) {}
        try { await window.storage.set("hsse-seed-version", SEED_VERSION, false); } catch (e) {}
      }
      setLoaded(true);
    })();
  }, []);

  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get("workhours-access-log", true);
        if (res && res.value) {
          const parsed = JSON.parse(res.value);
          if (Array.isArray(parsed)) setAccessLog(parsed);
        }
      } catch (e) {}
    })();
  }, []);

  const logAccessAttempt = async (usernameTried, success) => {
    const entry = { ts: new Date().toISOString(), username: usernameTried, success };
    setAccessLog((prev) => {
      const next = [entry, ...prev].slice(0, 200);
      (async () => {
        try { await window.storage.set("workhours-access-log", JSON.stringify(next), true); } catch (e) {}
      })();
      return next;
    });
  };

  const requestAdminAction = (action) => {
    if (adminUnlocked) { runAdminAction(action); return; }
    setPendingAction(action);
    setShowAccessGate(true);
  };

  const runAdminAction = (action) => {
    if (action === "logmonth") setShowEntry(true);
    if (action === "upload") fileInputRef.current?.click();
    if (action === "export") exportJSON();
  };

  useEffect(() => {
    if (!loaded) return;
    (async () => {
      try { await window.storage.set("workhours-months", JSON.stringify(months), false); } catch (e) { console.error(e); }
    })();
  }, [months, loaded]);

  useEffect(() => {
    if (!loaded) return;
    (async () => {
      try { await window.storage.set("workhours-incidents", JSON.stringify(incidents), false); } catch (e) { console.error(e); }
    })();
  }, [incidents, loaded]);

  useEffect(() => {
    if (!loaded) return;
    (async () => {
      try { await window.storage.set("workhours-custom-vendors", JSON.stringify(customVendors), false); } catch (e) { console.error(e); }
    })();
  }, [customVendors, loaded]);

  useEffect(() => {
    const id = setInterval(() => setTick((t) => t + 1), 2600);
    return () => clearInterval(id);
  }, []);

  const showToast = (msg, kind = "ok") => {
    setToast({ msg, kind });
    setTimeout(() => setToast(null), 3200);
  };

  /* ---- derived data ---- */

  const sorted = useMemo(() => [...months].sort((a, b) => a.month.localeCompare(b.month)), [months]);

  const effectiveSiteCategories = useMemo(() => [...SITE_CATEGORIES, ...customVendors.site], [customVendors]);
  const effectiveOffsiteCategories = useMemo(() => [...OFFSITE_CATEGORIES, ...customVendors.offsite], [customVendors]);

  const totals = useMemo(() => {
    let ptdSite = 0, ptdOffsite = 0;
    const byYear = {};
    sorted.forEach((m) => {
      const staff = m.staffPermanent || { site: 0, offsite: 0 };
      const siteSum = sum(m.site) + (Number(staff.site) || 0);
      const offsiteSum = sum(m.offsite) + (Number(staff.offsite) || 0);
      ptdSite += siteSum;
      ptdOffsite += offsiteSum;
      const yr = m.month.slice(0, 4);
      byYear[yr] = (byYear[yr] || 0) + siteSum + offsiteSum;
    });
    return { ptdSite, ptdOffsite, ptd: ptdSite + ptdOffsite, byYear };
  }, [sorted]);

  const indicatorTotals = useMemo(() => {
    const result = {};
    [...LAGGING_METRICS, ...LEADING_METRICS, ...WASTE_METRICS].forEach((m) => {
      const baseline = SEED_PTD_SUMMARY[m.key] || { site: 0, offsite: 0 };
      let site = Number(baseline.site) || 0;
      let offsite = Number(baseline.offsite) || 0;
      sorted.forEach((month) => {
        if (month.month > BASELINE_CUTOFF && month.indicators && month.indicators[m.key]) {
          site += Number(month.indicators[m.key].site) || 0;
          offsite += Number(month.indicators[m.key].offsite) || 0;
        }
      });
      result[m.key] = { site, offsite, ptd: site + offsite };
    });
    return result;
  }, [sorted]);

  const trendData = useMemo(
    () =>
      sorted.map((m) => {
        const staff = m.staffPermanent || { site: 0, offsite: 0 };
        return { month: monthLabel(m.month), Site: Math.round(sum(m.site) + (Number(staff.site) || 0)), Offsite: Math.round(sum(m.offsite) + (Number(staff.offsite) || 0)) };
      }),
    [sorted]
  );

  const latest = sorted[sorted.length - 1];
  const previous = sorted[sorted.length - 2];
  const monthTrend = useMemo(() => {
    if (!latest || !previous || !previous.total) return null;
    const pct = ((latest.total - previous.total) / previous.total) * 100;
    return { pct, up: pct >= 0 };
  }, [latest, previous]);

  const allContractorTotals = useMemo(() => {
    const totalsMap = {};
    sorted.forEach((m) => {
      effectiveSiteCategories.forEach((c) => {
        const key = c.trim();
        if (!totalsMap[key]) totalsMap[key] = { name: key, site: 0, offsite: 0 };
        totalsMap[key].site += Number(m.site[c]) || 0;
      });
      effectiveOffsiteCategories.forEach((c) => {
        const key = c.trim();
        if (!totalsMap[key]) totalsMap[key] = { name: key, site: 0, offsite: 0 };
        totalsMap[key].offsite += Number(m.offsite[c]) || 0;
      });
    });
    return Object.values(totalsMap)
      .map((d) => ({ name: d.name, site: Math.round(d.site), offsite: Math.round(d.offsite), hours: Math.round(d.site + d.offsite) }))
      .filter((d) => d.hours > 0);
  }, [sorted, effectiveSiteCategories, effectiveOffsiteCategories]);

  const contractorTotals = useMemo(
    () => [...allContractorTotals].sort((a, b) => b.hours - a.hours).slice(0, 8),
    [allContractorTotals]
  );

  const [contractorSort, setContractorSort] = useState({ key: "hours", dir: "desc" });
  const sortedContractorTotals = useMemo(() => {
    const { key, dir } = contractorSort;
    const list = [...allContractorTotals];
    list.sort((a, b) => {
      const va = key === "name" ? a.name.toLowerCase() : a[key];
      const vb = key === "name" ? b.name.toLowerCase() : b[key];
      if (va < vb) return dir === "asc" ? -1 : 1;
      if (va > vb) return dir === "asc" ? 1 : -1;
      return 0;
    });
    return list;
  }, [allContractorTotals, contractorSort]);
  const toggleContractorSort = (key) => setContractorSort((s) => (s.key === key ? { key, dir: s.dir === "asc" ? "desc" : "asc" } : { key, dir: "desc" }));

  const [monthSort, setMonthSort] = useState({ key: "month", dir: "desc" });
  const sortedMonths = useMemo(() => {
    const { key, dir } = monthSort;
    const list = sorted.map((m) => {
      const staff = m.staffPermanent || { site: 0, offsite: 0 };
      const site = Math.round(sum(m.site));
      const offsite = Math.round(sum(m.offsite));
      const staffTotal = Math.round((Number(staff.site) || 0) + (Number(staff.offsite) || 0));
      return { month: m.month, site, offsite, staff: staffTotal, total: site + offsite + staffTotal };
    });
    list.sort((a, b) => {
      const va = a[key], vb = b[key];
      if (va < vb) return dir === "asc" ? -1 : 1;
      if (va > vb) return dir === "asc" ? 1 : -1;
      return 0;
    });
    return list;
  }, [sorted, monthSort]);
  const toggleMonthSort = (key) => setMonthSort((s) => (s.key === key ? { key, dir: s.dir === "asc" ? "desc" : "asc" } : { key, dir: key === "month" ? "desc" : "desc" }));

  const staffVsContractor = useMemo(
    () =>
      sorted.map((m) => {
        const staff = m.staffPermanent || { site: 0, offsite: 0 };
        return {
          month: monthLabel(m.month),
          Staff: Math.round((Number(staff.site) || 0) + (Number(staff.offsite) || 0)),
          Contractor: Math.round(sum(m.site) + sum(m.offsite)),
        };
      }),
    [sorted]
  );

  const currentMonthStaff = useMemo(() => {
    if (!latest) return { site: 0, offsite: 0, total: 0, contractorTotal: 0, month: null };
    const staff = latest.staffPermanent || { site: 0, offsite: 0 };
    return {
      site: Number(staff.site) || 0,
      offsite: Number(staff.offsite) || 0,
      total: (Number(staff.site) || 0) + (Number(staff.offsite) || 0),
      contractorTotal: sum(latest.site) + sum(latest.offsite),
      month: latest.month,
    };
  }, [latest]);

  const vendorCount = useMemo(() => {
    const totalsMap = {};
    sorted.forEach((m) => {
      effectiveSiteCategories.forEach((c) => { totalsMap[c.trim()] = (totalsMap[c.trim()] || 0) + (Number(m.site[c]) || 0); });
      effectiveOffsiteCategories.forEach((c) => { totalsMap[c.trim()] = (totalsMap[c.trim()] || 0) + (Number(m.offsite[c]) || 0); });
    });
    const active = Object.values(totalsMap).filter((v) => v > 0).length;
    const roster = Object.keys(totalsMap).length;
    return { active, roster };
  }, [sorted, effectiveSiteCategories, effectiveOffsiteCategories]);

  const monthsSinceStart = useMemo(() => {
    const [sy, sm] = PROJECT_START.split("-").map(Number);
    const now = new Date();
    return (now.getFullYear() - sy) * 12 + (now.getMonth() + 1 - sm) + 1;
  }, []);

  const daysSinceStart = useMemo(() => {
    const start = new Date(2025, 3, 9); // 09 Apr 2025
    const now = new Date();
    return Math.max(0, Math.round((now.setHours(0, 0, 0, 0) - start.setHours(0, 0, 0, 0)) / 86400000));
  }, []);

  const splitData = [
    { name: "Site", value: Math.round(totals.ptdSite) },
    { name: "Offsite / Office", value: Math.round(totals.ptdOffsite) },
  ];

  const incidentStats = useMemo(() => {
    const byType = {};
    const byTypeYTD = {};
    const byVendor = {};
    const byVendorType = {};
    const types = new Set();
    const currentYear = String(new Date().getFullYear());
    let totalYTD = 0;
    incidents.forEach((i) => {
      byType[i.type] = (byType[i.type] || 0) + 1;
      byVendor[i.vendor] = (byVendor[i.vendor] || 0) + 1;
      types.add(i.type);
      if (!byVendorType[i.vendor]) byVendorType[i.vendor] = {};
      byVendorType[i.vendor][i.type] = (byVendorType[i.vendor][i.type] || 0) + 1;
      if ((i.date || "").startsWith(currentYear)) {
        byTypeYTD[i.type] = (byTypeYTD[i.type] || 0) + 1;
        totalYTD += 1;
      }
    });
    const vendorChart = Object.entries(byVendor)
      .map(([name, count]) => ({ name, count, ...byVendorType[name] }))
      .sort((a, b) => b.count - a.count);
    return { byType, byTypeYTD, vendorChart, total: incidents.length, totalYTD, types: Array.from(types) };
  }, [incidents]);

  /* ---- upload handling ---- */

  const handleFile = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
        const sheetName = wb.SheetNames.find((n) => /site-month/i.test(n)) || wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const grid = XLSX.utils.sheet_to_json(ws, { header: "A", raw: true, defval: null });

        let headerRowIdx = grid.findIndex((r) => (r.A === "PTD" && r.B === "SITE") || (r.D === "EcoOils" && r.A === "SITE"));
        let siteColLetters, offsiteColLetters, headerRow;

        if (grid.findIndex((r) => r.D === "EcoOils" && r.A === "SITE") !== -1) {
          // new master template layout
          headerRowIdx = grid.findIndex((r) => r.D === "EcoOils" && r.A === "SITE");
          siteColLetters = ["D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T"];
          offsiteColLetters = ["U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI"];
        } else if (headerRowIdx !== -1) {
          // legacy template layout
          siteColLetters = ["E","F","G","H","I","J","K","L","M","N","O","P","Q"];
          offsiteColLetters = ["S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL"];
        } else {
          throw new Error("Could not find a recognizable header row — is this a Mentari Workhours/HSE Data template?");
        }
        headerRow = grid[headerRowIdx];
        const siteNames = siteColLetters.map((c) => headerRow[c]);
        const offsiteNames = offsiteColLetters.map((c) => headerRow[c]);

        // Scan every row after the header for a real date — don't assume a fixed
        // row offset, since blank spacer rows can get silently compacted out of
        // the parsed grid and throw a fixed offset off.
        const parsedMonths = [];
        for (let i = headerRowIdx + 1; i < grid.length; i++) {
          const row = grid[i];
          if (!row || !row.A) continue;
          const dateVal = row.A ?? row.B;
          let d;
          if (dateVal instanceof Date) d = dateVal;
          else continue;
          if (!d || isNaN(d.getTime())) continue;

          const monthKey = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
          const siteObj = {};
          siteColLetters.forEach((c, idx) => (siteObj[siteNames[idx]] = Number(row[c]) || 0));
          const offsiteObj = {};
          offsiteColLetters.forEach((c, idx) => (offsiteObj[offsiteNames[idx]] = Number(row[c]) || 0));
          // EcoOils is staff, not a contractor — pull it out of the vendor totals.
          const staffSite = siteObj["EcoOils"] || 0;
          const staffOffsite = offsiteObj["EcoOils"] || 0;
          delete siteObj["EcoOils"];
          delete offsiteObj["EcoOils"];
          const total = sum(siteObj) + sum(offsiteObj) + staffSite + staffOffsite;
          if (total === 0 && d.getFullYear() >= new Date().getFullYear() + 1) continue;
          parsedMonths.push({ month: monthKey, total, site: siteObj, offsite: offsiteObj, staffPermanent: { site: staffSite, offsite: staffOffsite } });
        }

        if (!parsedMonths.length) throw new Error("No monthly rows found in this file.");

        setMonths((prev) => {
          const map = new Map(prev.map((m) => [m.month, m]));
          parsedMonths.forEach((m) => map.set(m.month, m));
          return Array.from(map.values());
        });
        showToast(`Imported ${parsedMonths.length} month(s) from ${file.name}`, "ok");
      } catch (err) {
        showToast(err.message || "Could not parse this file", "err");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  /* ---- export ---- */

  const exportJSON = () => {
    const blob = new Blob([JSON.stringify({ site_categories: effectiveSiteCategories, offsite_categories: effectiveOffsiteCategories, months: sorted, indicator_totals: indicatorTotals, weekly: WEEKLY, incidents }, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `mentari-hsse-snapshot-${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
    showToast("Snapshot downloaded — send it back in chat to generate a PPT", "ok");
  };

  const digits = fmt(totals.ptd).padStart(11, "0").split("");

  return (
    <div style={{ background: PALETTE.bg, minHeight: "100vh", color: PALETTE.text, fontFamily: "'Inter', -apple-system, sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;700&display=swap');
        @import url('https://fonts.googleapis.com/css2?family=Digital+Numbers&display=swap');
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-thumb { background: ${PALETTE.border}; border-radius: 4px; }
        .mono { font-family: 'JetBrains Mono', ui-monospace, 'SF Mono', Menlo, monospace; }
        .heading { font-family: 'Space Grotesk', 'Inter', -apple-system, sans-serif; letter-spacing: 0.1px; }
        .digital-clock { font-family: 'Digital Numbers', 'JetBrains Mono', monospace; }
        input[type=number]::-webkit-inner-spin-button { opacity: 1; }
        button:focus-visible, input:focus-visible { outline: 2px solid ${PALETTE.amber}; outline-offset: 2px; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: translateY(0); } }
        @keyframes pulseDot { 0%,100% { opacity: 1; } 50% { opacity: 0.35; } }
      `}</style>

      {/* Header */}
      <div style={{ borderBottom: `1px solid ${PALETTE.border}` }}>
        <div style={{ padding: "18px 28px 14px", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 16 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
            <img src={LOGO_DATA_URI} alt="EcoOils Mentari" style={{ height: 34, width: "auto", display: "block" }} />
            <div style={{ width: 1, alignSelf: "stretch", background: PALETTE.border, margin: "2px 0" }} />
            <div>
              <div className="heading" style={{ fontSize: 17, fontWeight: 700, letterSpacing: 0.2 }}>HSSE Performance</div>
              <div className="mono" style={{ fontSize: 11.5, color: PALETTE.textDim, marginTop: 2 }}>Data as of {DATA_AS_OF} · Build {SEED_VERSION}</div>
            </div>
          </div>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
            <button onClick={() => requestAdminAction("logmonth")} style={btnStyle(PALETTE.amber, PALETTE.bg)}>
              <PlusCircle size={15} /> Log month
            </button>
            {adminUnlocked && (
              <>
                <button onClick={() => requestAdminAction("upload")} style={btnStyle("transparent", PALETTE.text, PALETTE.border)}>
                  <Upload size={15} /> Upload file
                </button>
                <button onClick={() => requestAdminAction("export")} style={btnStyle("transparent", PALETTE.text, PALETTE.border)}>
                  <Download size={15} /> Export
                </button>
                <button onClick={() => setShowAccessLog(true)} style={{ background: "none", border: "none", color: PALETTE.textDim, cursor: "pointer", display: "flex", alignItems: "center", gap: 4, fontSize: 11.5 }}>
                  <Lock size={12} /> Access log
                </button>
              </>
            )}
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={(e) => { if (e.target.files[0]) handleFile(e.target.files[0]); e.target.value = ""; }} />
          </div>
        </div>

        {/* Project / Location context bar */}
        <div style={{ padding: "10px 28px", display: "flex", alignItems: "center", gap: 22, flexWrap: "wrap", background: PALETTE.panelAlt, borderTop: `1px solid ${PALETTE.border}` }}>
          <ContextSelect label="Project" value={project} onChange={setProject} options={PROJECTS.map((p) => ({ value: p.id, label: p.name }))} />
          <ContextSelect label="Location" value={location} onChange={setLocation} options={(PROJECTS.find((p) => p.id === project)?.locations || []).map((l) => ({ value: l, label: l }))} />
          <div className="mono" style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 8, fontSize: 11, color: PALETTE.textDim, background: "#000", border: `1px solid ${PALETTE.amberDim}`, borderRadius: 20, padding: "4px 12px 4px 6px" }}>
            <span style={{ width: 6, height: 6, borderRadius: "50%", background: PALETTE.amber, animation: "pulseDot 2.6s ease-in-out infinite", flexShrink: 0 }} />
            <span>Site construction commenced 09 Apr 2025</span>
            <span style={{ color: PALETTE.amber, fontWeight: 700 }}>· Day {fmt(daysSinceStart)} · Month {monthsSinceStart}</span>
          </div>
        </div>
      </div>

      <div style={{ maxWidth: 1200, margin: "0 auto", padding: "18px" }}>
        {/* Work Hours — consolidated section */}
        <Panel
          title="Work Hours"
        >
          {/* Signature: LED-style cumulative counter */}
          <div style={{ position: "relative", overflow: "hidden", borderRadius: 10, background: "#0B1013", border: `1px solid ${PALETTE.border}`, padding: "18px 20px", marginBottom: 14 }}>
            <div style={{ position: "absolute", inset: 0, background: `radial-gradient(circle at 15% 20%, ${PALETTE.amberDim}22, transparent 55%)`, pointerEvents: "none" }} />
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", flexWrap: "wrap", gap: 16, position: "relative" }}>
              <div>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                  <span style={{ width: 6, height: 6, borderRadius: "50%", background: PALETTE.amber, animation: "pulseDot 2.6s ease-in-out infinite" }} />
                  <span style={{ fontSize: 10.5, textTransform: "uppercase", letterSpacing: 1.3, color: PALETTE.textDim, fontWeight: 600 }}>Project-to-date man-hours</span>
                </div>
                <div className="digital-clock" style={{ display: "flex", gap: 3 }}>
                  {digits.map((d, i) => {
                    const isSpacer = d === "," || d === ".";
                    return (
                      <span key={i} style={{ display: "inline-block", minWidth: isSpacer ? 8 : 24, textAlign: "center", fontSize: 32, fontWeight: 700, color: isSpacer ? PALETTE.textDim : PALETTE.amber, background: isSpacer ? "transparent" : "#000", border: isSpacer ? "none" : `1px solid ${PALETTE.amberDim}`, borderRadius: 4, padding: isSpacer ? 0 : "1px 0", textShadow: isSpacer ? "none" : `0 0 12px ${PALETTE.amber}55` }}>
                        {d}
                      </span>
                    );
                  })}
                </div>
                <div style={{ fontSize: 11.5, color: PALETTE.textDim, marginTop: 8 }}>
                  Site <b style={{ color: PALETTE.text }}>{fmt(totals.ptdSite)}</b> hrs &nbsp;·&nbsp; Offsite <b style={{ color: PALETTE.text }}>{fmt(totals.ptdOffsite)}</b> hrs
                </div>
              </div>
              <div style={{ display: "flex", gap: 20, flexWrap: "wrap" }}>
                {Object.entries(totals.byYear).sort().map(([yr, val]) => (
                  <div key={yr}>
                    <div style={{ fontSize: 10.5, color: PALETTE.textDim, textTransform: "uppercase", letterSpacing: 0.8 }}>{yr} YTD</div>
                    <div className="mono" style={{ fontSize: 19, fontWeight: 700, color: PALETTE.text }}>{fmt(val)}</div>
                  </div>
                ))}
                {latest && (
                  <div>
                    <div style={{ fontSize: 10.5, color: PALETTE.textDim, textTransform: "uppercase", letterSpacing: 0.8 }}>Latest ({monthLabel(latest.month)})</div>
                    <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                      <div className="mono" style={{ fontSize: 19, fontWeight: 700, color: PALETTE.teal }}>{fmt(latest.total)}</div>
                      {monthTrend && (
                        <span style={{ display: "flex", alignItems: "center", gap: 3, fontSize: 11, fontWeight: 700, color: monthTrend.up ? "#4ADE80" : "#F87171" }}>
                          {monthTrend.up ? <TrendingUp size={13} /> : <TrendingDown size={13} />}
                          {Math.abs(monthTrend.pct).toFixed(0)}%
                        </span>
                      )}
                    </div>
                  </div>
                )}
              </div>
            </div>
          </div>

          {/* Compact charts row */}
          <div style={{ display: "grid", gridTemplateColumns: "1.6fr 1fr", gap: 12 }}>
            <div>
              <div style={{ fontSize: 11, color: PALETTE.textDim, marginBottom: 6 }}>Monthly hours — site vs. offsite</div>
              <ResponsiveContainer width="100%" height={160}>
                <AreaChart data={trendData} margin={{ top: 4, right: 8, left: -18, bottom: 0 }}>
                  <defs>
                    <linearGradient id="siteGrad" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="0%" stopColor={SITE_COLOR} stopOpacity={0.55} />
                      <stop offset="100%" stopColor={SITE_COLOR} stopOpacity={0.02} />
                    </linearGradient>
                    <linearGradient id="offsiteGrad" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="0%" stopColor={OFFSITE_COLOR} stopOpacity={0.5} />
                      <stop offset="100%" stopColor={OFFSITE_COLOR} stopOpacity={0.02} />
                    </linearGradient>
                  </defs>
                  <CartesianGrid strokeDasharray="3 3" stroke={PALETTE.border} vertical={false} />
                  <XAxis dataKey="month" stroke={PALETTE.textDim} fontSize={10} tickLine={false} axisLine={{ stroke: PALETTE.border }} interval={2} />
                  <YAxis stroke={PALETTE.textDim} fontSize={10} tickLine={false} axisLine={false} tickFormatter={(v) => `${v / 1000}k`} />
                  <Tooltip contentStyle={tooltipStyle} formatter={(v) => fmt(v)} />
                  <Area type="monotone" dataKey="Site" stackId="1" stroke={SITE_COLOR} fill="url(#siteGrad)" strokeWidth={2} />
                  <Area type="monotone" dataKey="Offsite" stackId="1" stroke={OFFSITE_COLOR} fill="url(#offsiteGrad)" strokeWidth={2} />
                </AreaChart>
              </ResponsiveContainer>
            </div>
            <div>
              <div style={{ fontSize: 11, color: PALETTE.textDim, marginBottom: 6 }}>PTD split</div>
              <ResponsiveContainer width="100%" height={160}>
                <PieChart>
                  <Pie data={splitData} dataKey="value" nameKey="name" innerRadius={38} outerRadius={62} paddingAngle={3}>
                    <Cell fill={SITE_COLOR} />
                    <Cell fill={OFFSITE_COLOR} />
                  </Pie>
                  <Tooltip contentStyle={tooltipStyle} formatter={(v) => fmt(v)} />
                  <Legend wrapperStyle={{ fontSize: 10.5, color: PALETTE.textDim }} />
                </PieChart>
              </ResponsiveContainer>
            </div>
          </div>

          <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 10 }}>
            <button onClick={() => setShowWorkHoursDetail((v) => !v)} style={{ background: "none", border: "none", color: PALETTE.textDim, cursor: "pointer", display: "flex", alignItems: "center", gap: 5, fontSize: 12 }}>
              {showWorkHoursDetail ? "Hide details" : "Show details"}
              {showWorkHoursDetail ? <ChevronUp size={15} /> : <ChevronDown size={15} />}
            </button>
          </div>

          {showWorkHoursDetail && (
            <div style={{ marginTop: 16, paddingTop: 16, borderTop: `1px solid ${PALETTE.border}`, display: "flex", flexDirection: "column", gap: 16 }}>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 10 }}>
                <KpiCard icon={<Clock size={16} color={OFFSITE_COLOR} />} label="Offsite share of PTD" value={`${((totals.ptdOffsite / (totals.ptd || 1)) * 100).toFixed(0)}%`} sub={`${fmt(totals.ptdOffsite)} hrs offsite/office`} />
                <KpiCard icon={<TrendingUp size={16} color={PALETTE.amber} />} label="Months tracked" value={sorted.length} sub={`Site commenced 09 Apr '25 · ${monthsSinceStart} mo elapsed`} />
                <KpiCard icon={<MapPin size={16} color={PALETTE.teal} />} label="Vendors" value={vendorCount.active} sub={`${vendorCount.active} active of ${vendorCount.roster} on roster`} />
              </div>

              <div>
                <SectionLabel icon={<TrendingUp size={13} color={PALETTE.amber} />} text="Top contractors — cumulative hours (PTD)" />
                <ResponsiveContainer width="100%" height={240}>
                  <BarChart data={contractorTotals} layout="vertical" margin={{ top: 6, right: 24, left: 10, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke={PALETTE.border} horizontal={false} />
                    <XAxis type="number" stroke={PALETTE.textDim} fontSize={11} tickFormatter={(v) => `${v / 1000}k`} tickLine={false} axisLine={{ stroke: PALETTE.border }} />
                    <YAxis type="category" dataKey="name" stroke={PALETTE.textDim} fontSize={12} width={120} tickLine={false} axisLine={false} />
                    <Tooltip contentStyle={tooltipStyle} formatter={(v) => fmt(v)} />
                    <Legend wrapperStyle={{ fontSize: 12, color: PALETTE.textDim }} />
                    <Bar dataKey="site" name="Site" stackId="ptd" fill={SITE_COLOR} radius={[0, 0, 0, 0]} />
                    <Bar dataKey="offsite" name="Offsite" stackId="ptd" fill={OFFSITE_COLOR} radius={[0, 4, 4, 0]} />
                  </BarChart>
                </ResponsiveContainer>
                <div style={{ overflowX: "auto", marginTop: 12 }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                    <thead>
                      <tr style={{ textAlign: "left", color: PALETTE.textDim, fontSize: 11, textTransform: "uppercase", letterSpacing: 0.5 }}>
                        <SortHeader label="Vendor" active={contractorSort.key === "name"} dir={contractorSort.dir} onClick={() => toggleContractorSort("name")} />
                        <SortHeader label="Site hrs" active={contractorSort.key === "site"} dir={contractorSort.dir} onClick={() => toggleContractorSort("site")} align="right" />
                        <SortHeader label="Offsite hrs" active={contractorSort.key === "offsite"} dir={contractorSort.dir} onClick={() => toggleContractorSort("offsite")} align="right" />
                        <SortHeader label="Total hrs" active={contractorSort.key === "hours"} dir={contractorSort.dir} onClick={() => toggleContractorSort("hours")} align="right" />
                      </tr>
                    </thead>
                    <tbody>
                      {sortedContractorTotals.map((c) => (
                        <tr key={c.name} style={{ borderTop: `1px solid ${PALETTE.border}` }}>
                          <td style={{ ...thtd, fontWeight: 600 }}>{vendorLabel(c.name)}</td>
                          <td className="mono" style={{ ...thtd, textAlign: "right", color: SITE_COLOR }}>{fmt(c.site)}</td>
                          <td className="mono" style={{ ...thtd, textAlign: "right", color: OFFSITE_COLOR }}>{fmt(c.offsite)}</td>
                          <td className="mono" style={{ ...thtd, textAlign: "right", fontWeight: 700 }}>{fmt(c.hours)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              <div>
                <SectionLabel icon={<Users size={13} color={PALETTE.purple} />} text="EcoOils Staff Permanent vs. Contractor" />
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 10, marginBottom: 12 }}>
                  <MiniStat label={`Staff — Site (${currentMonthStaff.month ? monthLabel(currentMonthStaff.month) : "—"})`} value={fmt(currentMonthStaff.site)} color={PALETTE.purple} />
                  <MiniStat label={`Staff — Offsite (${currentMonthStaff.month ? monthLabel(currentMonthStaff.month) : "—"})`} value={fmt(currentMonthStaff.offsite)} color={PALETTE.purple} />
                  <MiniStat label="Staff total, this month" value={fmt(currentMonthStaff.total)} color={PALETTE.purple} />
                  <MiniStat label="Contractor total, this month" value={fmt(currentMonthStaff.contractorTotal)} color={PALETTE.teal} />
                </div>
                <ResponsiveContainer width="100%" height={180}>
                  <BarChart data={staffVsContractor} margin={{ top: 6, right: 12, left: -14, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke={PALETTE.border} vertical={false} />
                    <XAxis dataKey="month" stroke={PALETTE.textDim} fontSize={11} tickLine={false} axisLine={{ stroke: PALETTE.border }} />
                    <YAxis stroke={PALETTE.textDim} fontSize={11} tickLine={false} axisLine={false} tickFormatter={(v) => `${v / 1000}k`} />
                    <Tooltip contentStyle={tooltipStyle} formatter={(v) => fmt(v)} />
                    <Legend wrapperStyle={{ fontSize: 12, color: PALETTE.textDim }} />
                    <Bar dataKey="Staff" fill={PALETTE.purple} radius={[3, 3, 0, 0]} />
                    <Bar dataKey="Contractor" fill={PALETTE.teal} radius={[3, 3, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>

              <div>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                  <SectionLabel icon={<ClipboardCheck size={13} color={PALETTE.amber} />} text={`Monthly log (${sorted.length} entries)`} />
                </div>
                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                    <thead>
                      <tr style={{ textAlign: "left", color: PALETTE.textDim, fontSize: 11, textTransform: "uppercase", letterSpacing: 0.5 }}>
                        <SortHeader label="Month" active={monthSort.key === "month"} dir={monthSort.dir} onClick={() => toggleMonthSort("month")} />
                        <SortHeader label="Site hrs" active={monthSort.key === "site"} dir={monthSort.dir} onClick={() => toggleMonthSort("site")} align="right" />
                        <SortHeader label="Offsite hrs" active={monthSort.key === "offsite"} dir={monthSort.dir} onClick={() => toggleMonthSort("offsite")} align="right" />
                        <SortHeader label="Staff (Perm.)" active={monthSort.key === "staff"} dir={monthSort.dir} onClick={() => toggleMonthSort("staff")} align="right" />
                        <SortHeader label="Total" active={monthSort.key === "total"} dir={monthSort.dir} onClick={() => toggleMonthSort("total")} align="right" />
                      </tr>
                    </thead>
                    <tbody>
                      {sortedMonths.map((m) => (
                        <tr key={m.month} style={{ borderTop: `1px solid ${PALETTE.border}` }}>
                          <td style={{ ...thtd, fontWeight: 600 }}>{monthLabel(m.month)}</td>
                          <td className="mono" style={{ ...thtd, textAlign: "right", color: SITE_COLOR }}>{fmt(m.site)}</td>
                          <td className="mono" style={{ ...thtd, textAlign: "right", color: OFFSITE_COLOR }}>{fmt(m.offsite)}</td>
                          <td className="mono" style={{ ...thtd, textAlign: "right", color: PALETTE.purple }}>{fmt(m.staff)}</td>
                          <td className="mono" style={{ ...thtd, textAlign: "right", fontWeight: 700 }}>{fmt(m.total)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}
        </Panel>

        {/* HSSE performance scorecard */}
        <Panel
          title="HSSE Performance Scorecard — PTD"
          style={{ marginTop: 12 }}
        >
          <div style={{ overflowX: "auto", marginBottom: 14, paddingBottom: 2 }}>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(7, minmax(96px, 1fr))", gap: 8, minWidth: 720 }}>
              <BigStat label="TRCF" value={fmt2(indicatorTotals.trcf.ptd)} color={PALETTE.teal} />
              <BigStat label="LTIF" value={fmt2(indicatorTotals.ltif.ptd)} color={PALETTE.teal} />
              <BigStat label="LSR" value={fmt(indicatorTotals.lsr.ptd)} color={INCIDENT_TYPE_COLORS.LSR} />
              <BigStat label="First Aid Case" value={fmt(indicatorTotals.fac.ptd)} color={INCIDENT_TYPE_COLORS["First Aid Case"]} />
              <BigStat label="Near Miss" value={fmt(indicatorTotals.nearMiss.ptd)} color={INCIDENT_TYPE_COLORS["Near Miss"]} />
              <BigStat label="Property Damage" value={fmt(indicatorTotals.equipDamage.ptd)} color={INCIDENT_TYPE_COLORS["Property Damage"]} />
              <BigStat label="Security" value={fmt(indicatorTotals.security.ptd)} color={INCIDENT_TYPE_COLORS.Security} />
            </div>
          </div>

          {/* Leading / Lagging visual header */}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, background: "#2A15131a", border: `1px solid ${PALETTE.red}55`, borderRadius: 8, padding: "8px 12px" }}>
              <ArrowLeft size={16} color={PALETTE.red} style={{ flexShrink: 0 }} />
              <div>
                <div className="heading" style={{ fontSize: 12.5, fontWeight: 700, color: PALETTE.red }}>LAGGING</div>
                <div style={{ fontSize: 10, color: PALETTE.textDim }}>Analyze past performance</div>
              </div>
            </div>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "flex-end", gap: 8, background: "#0F2A1a1a", border: `1px solid ${PALETTE.teal}55`, borderRadius: 8, padding: "8px 12px" }}>
              <div style={{ textAlign: "right" }}>
                <div className="heading" style={{ fontSize: 12.5, fontWeight: 700, color: PALETTE.teal }}>LEADING</div>
                <div style={{ fontSize: 10, color: PALETTE.textDim }}>Influence future performance</div>
              </div>
              <ArrowRight size={16} color={PALETTE.teal} style={{ flexShrink: 0 }} />
            </div>
          </div>

          <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 10 }}>
            <button onClick={() => setShowScorecardDetail((v) => !v)} style={{ background: "none", border: "none", color: PALETTE.textDim, cursor: "pointer", display: "flex", alignItems: "center", gap: 5, fontSize: 12 }}>
              {showScorecardDetail ? "Hide details" : "Show details"}
              {showScorecardDetail ? <ChevronUp size={15} /> : <ChevronDown size={15} />}
            </button>
          </div>

          {showScorecardDetail && (
            <div style={{ marginTop: 8, paddingTop: 16, borderTop: `1px solid ${PALETTE.border}`, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 24 }}>
              <div>
                <SectionLabel icon={<Siren size={13} color={PALETTE.red} />} text="Lagging indicators" />
                <ScoreTable rows={LAGGING_METRICS} data={indicatorTotals} accent={PALETTE.red} />
              </div>
              <div>
                <SectionLabel icon={<ClipboardCheck size={13} color={PALETTE.teal} />} text="Leading indicators" />
                <ScoreTable rows={LEADING_METRICS} data={indicatorTotals} accent={PALETTE.teal} />
                <div style={{ marginTop: 18 }}>
                  <SectionLabel icon={<Trash2 size={13} color={PALETTE.textDim} />} text="Waste disposal" />
                  <ScoreTable rows={WASTE_METRICS} data={indicatorTotals} accent={PALETTE.textDim} />
                </div>
              </div>
            </div>
          )}
        </Panel>

        {/* Incident & Lessons Learned Log */}
        <Panel title="Incidents PTD" style={{ marginTop: 14 }}>
          <div style={{ display: "grid", gridTemplateColumns: "1.5fr 1fr", gap: 20 }}>
            <ResponsiveContainer width="100%" height={Math.max(200, incidentStats.vendorChart.length * 34)}>
              <BarChart data={incidentStats.vendorChart} layout="vertical" margin={{ top: 0, right: 24, left: 10, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={PALETTE.border} horizontal={false} />
                <XAxis type="number" stroke={PALETTE.textDim} fontSize={11} allowDecimals={false} tickLine={false} axisLine={{ stroke: PALETTE.border }} />
                <YAxis type="category" dataKey="name" stroke={PALETTE.textDim} fontSize={12} width={70} tickLine={false} axisLine={false} />
                <Tooltip contentStyle={tooltipStyle} />
                <Legend wrapperStyle={{ fontSize: 11, color: PALETTE.textDim }} />
                {incidentStats.types.map((type) => (
                  <Bar key={type} dataKey={type} name={type} stackId="incidents" fill={INCIDENT_TYPE_COLORS[type] || PALETTE.textDim} radius={[0, 0, 0, 0]} barSize={16} />
                ))}
              </BarChart>
            </ResponsiveContainer>

            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, minWidth: 260 }}>
                <thead>
                  <tr style={{ textAlign: "left", color: PALETTE.textDim, fontSize: 10.5, textTransform: "uppercase", letterSpacing: 0.5, background: PALETTE.panelAlt }}>
                    <th style={{ padding: "8px 10px", borderRadius: "8px 0 0 8px" }}>Type</th>
                    <th style={{ padding: "8px 10px", textAlign: "right" }}>PTD</th>
                    <th style={{ padding: "8px 10px", textAlign: "right", borderRadius: "0 8px 8px 0" }}>YTD ({new Date().getFullYear()})</th>
                  </tr>
                </thead>
                <tbody>
                  {incidentStats.types.map((type, i) => (
                    <tr key={type} style={{ borderTop: `1px solid ${PALETTE.border}`, background: i % 2 === 1 ? PALETTE.panelAlt : "transparent" }}>
                      <td style={{ padding: "8px 10px", display: "flex", alignItems: "center", gap: 8 }}>
                        <span style={{ width: 8, height: 8, borderRadius: "50%", background: INCIDENT_TYPE_COLORS[type] || PALETTE.textDim, flexShrink: 0 }} />
                        {type}
                      </td>
                      <td className="mono" style={{ padding: "8px 10px", textAlign: "right", fontWeight: 700, color: INCIDENT_TYPE_COLORS[type] || PALETTE.text }}>{incidentStats.byType[type] || 0}</td>
                      <td className="mono" style={{ padding: "8px 10px", textAlign: "right", color: PALETTE.textDim }}>{incidentStats.byTypeYTD[type] || 0}</td>
                    </tr>
                  ))}
                  <tr style={{ borderTop: `2px solid ${PALETTE.border}`, background: PALETTE.panelAlt }}>
                    <td style={{ padding: "8px 10px", fontWeight: 700, color: PALETTE.text }}>Total incidents</td>
                    <td className="mono" style={{ padding: "8px 10px", textAlign: "right", fontWeight: 700, color: PALETTE.text }}>{incidentStats.total}</td>
                    <td className="mono" style={{ padding: "8px 10px", textAlign: "right", fontWeight: 700, color: PALETTE.text }}>{incidentStats.totalYTD}</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        </Panel>
      </div>

      {showEntry && (
        <EntryModal
          onClose={() => setShowEntry(false)}
          onSave={(m) => { setMonths((prev) => { const map = new Map(prev.map((x) => [x.month, x])); map.set(m.month, m); return Array.from(map.values()); }); setShowEntry(false); showToast(`Saved ${monthLabel(m.month)}`, "ok"); }}
          existing={sorted}
          siteCategories={effectiveSiteCategories}
          offsiteCategories={effectiveOffsiteCategories}
          onAddVendor={(section, name) => {
            setCustomVendors((prev) => {
              if (prev[section].includes(name) || SITE_CATEGORIES.includes(name) || OFFSITE_CATEGORIES.includes(name)) return prev;
              return { ...prev, [section]: [...prev[section], name] };
            });
            showToast(`Added ${name} as a contractor`, "ok");
          }}
        />
      )}

      {showAccessGate && (
        <AccessGateModal
          onClose={() => { setShowAccessGate(false); setPendingAction(null); }}
          onSubmit={(username, password) => {
            const success = username.trim().toLowerCase() === ADMIN_USERNAME && password === ADMIN_PASSWORD;
            logAccessAttempt(username.trim(), success);
            if (success) {
              setAdminUnlocked(true);
              setShowAccessGate(false);
              const action = pendingAction;
              setPendingAction(null);
              runAdminAction(action);
              showToast("Access granted", "ok");
              return true;
            }
            return false;
          }}
        />
      )}

      {showAccessLog && (
        <AccessLogModal onClose={() => setShowAccessLog(false)} entries={accessLog} />
      )}

      {toast && (
        <div style={{ position: "fixed", bottom: 24, left: "50%", transform: "translateX(-50%)", background: toast.kind === "err" ? "#3A1E1A" : "#1C2A1E", border: `1px solid ${toast.kind === "err" ? PALETTE.red : "#4C8B5C"}`, color: PALETTE.text, padding: "10px 18px", borderRadius: 8, fontSize: 13, display: "flex", alignItems: "center", gap: 8, animation: "fadeIn 0.2s ease-out", zIndex: 100 }}>
          {toast.kind === "err" ? <AlertCircle size={15} color={PALETTE.red} /> : <Check size={15} color="#4C8B5C" />}
          {toast.msg}
        </div>
      )}
    </div>
  );
}

function KpiCard({ icon, label, value, sub }) {
  return (
    <div style={{ background: PALETTE.panel, border: `1px solid ${PALETTE.border}`, borderRadius: 9, padding: "11px 13px" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 5 }}>
        {icon}
        <span style={{ fontSize: 10.5, color: PALETTE.textDim, textTransform: "uppercase", letterSpacing: 0.5 }}>{label}</span>
      </div>
      <div style={{ fontSize: 18, fontWeight: 700 }}>{value}</div>
      <div style={{ fontSize: 10.5, color: PALETTE.textDim, marginTop: 2 }}>{sub}</div>
    </div>
  );
}

function BigStat({ label, value, color }) {
  return (
    <div style={{ background: PALETTE.panel, border: `1px solid ${PALETTE.border}`, borderRadius: 9, padding: "9px 11px" }}>
      <div style={{ fontSize: 9.5, color: PALETTE.textDim, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 4 }}>{label}</div>
      <div className="mono" style={{ fontSize: 19, fontWeight: 700, color }}>{value}</div>
    </div>
  );
}

function MiniStat({ label, value, color, big }) {
  return (
    <div style={{ background: PALETTE.panel, border: `1px solid ${PALETTE.border}`, borderRadius: 9, padding: "9px 11px" }}>
      <div style={{ fontSize: 9.5, color: PALETTE.textDim, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 4 }}>{label}</div>
      <div className="mono" style={{ fontSize: big ? 24 : 19, fontWeight: 700, color }}>{value}</div>
    </div>
  );
}

function SortHeader({ label, active, dir, onClick, align }) {
  return (
    <th style={{ padding: "9px 10px", textAlign: align || "left", cursor: "pointer", userSelect: "none", whiteSpace: "nowrap" }} onClick={onClick}>
      <span style={{ display: "inline-flex", alignItems: "center", gap: 4, justifyContent: align === "right" ? "flex-end" : "flex-start", color: active ? PALETTE.text : PALETTE.textDim }}>
        {label}
        {active ? (dir === "asc" ? <ChevronUp size={12} /> : <ChevronDown size={12} />) : <ChevronDown size={12} style={{ opacity: 0.25 }} />}
      </span>
    </th>
  );
}

function ContextSelect({ label, value, onChange, options }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
      <span style={{ fontSize: 10.5, color: PALETTE.textDim, textTransform: "uppercase", letterSpacing: 0.6 }}>{label}</span>
      <select
        value={value}
        onChange={(e) => onChange(e.target.value)}
        style={{ background: "transparent", border: `1px solid ${PALETTE.border}`, borderRadius: 6, color: PALETTE.text, fontSize: 12.5, fontWeight: 600, padding: "5px 8px", cursor: "pointer" }}
      >
        {options.map((o) => (<option key={o.value} value={o.value} style={{ background: PALETTE.panel }}>{o.label}</option>))}
      </select>
    </div>
  );
}

function SectionLabel({ icon, text }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 10 }}>
      {icon}
      <span className="heading" style={{ fontSize: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.6, color: PALETTE.text }}>{text}</span>
    </div>
  );
}

function ScoreTable({ rows, data, accent }) {
  return (
    <div style={{ border: `1px solid ${PALETTE.border}`, borderRadius: 10, overflow: "hidden" }}>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
        <thead>
          <tr style={{ textAlign: "left", color: PALETTE.textDim, fontSize: 10, textTransform: "uppercase", letterSpacing: 0.4, background: PALETTE.panelAlt }}>
            <th style={{ padding: "7px 10px" }}>Indicator</th>
            <th style={{ padding: "7px 10px", textAlign: "right" }}>PTD</th>
            <th style={{ padding: "7px 10px", textAlign: "right" }}>Site</th>
            <th style={{ padding: "7px 10px", textAlign: "right" }}>Offsite</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((m, i) => {
            const d = data[m.key] || { ptd: 0, site: 0, offsite: 0 };
            const f = (m.key === "trcf" || m.key === "ltif") ? fmt2 : fmt1;
            const nonzero = Number(d.ptd) > 0;
            return (
              <tr key={m.key} style={{ borderTop: `1px solid ${PALETTE.border}`, background: i % 2 === 1 ? "#FFFFFF06" : "transparent" }}>
                <td style={{ padding: "7px 10px", color: PALETTE.text, borderLeft: nonzero ? `2px solid ${accent || PALETTE.amber}` : "2px solid transparent" }}>{m.label}</td>
                <td className="mono" style={{ padding: "7px 10px", textAlign: "right", fontWeight: 700, color: nonzero ? (accent || PALETTE.text) : PALETTE.text }}>{f(d.ptd)}{m.unit ? ` ${m.unit}` : ""}</td>
                <td className="mono" style={{ padding: "7px 10px", textAlign: "right", color: PALETTE.textDim }}>{f(d.site)}</td>
                <td className="mono" style={{ padding: "7px 10px", textAlign: "right", color: PALETTE.textDim }}>{f(d.offsite)}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

function Panel({ title, children, style, right }) {
  return (
    <div style={{ background: PALETTE.panel, border: `1px solid ${PALETTE.border}`, borderRadius: 12, padding: "14px 16px", ...style }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
        <div className="heading" style={{ fontSize: 13, fontWeight: 700, color: PALETTE.text }}>{title}</div>
        {right}
      </div>
      {children}
    </div>
  );
}

const tooltipStyle = { background: "#000", border: `1px solid ${PALETTE.border}`, borderRadius: 8, fontSize: 12, color: PALETTE.text };
const thtd = { padding: "9px 10px" };
const btnStyle = (bg, color, border) => ({
  display: "flex", alignItems: "center", gap: 6, background: bg, color, border: border ? `1px solid ${border}` : "none",
  padding: "9px 14px", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer",
});

/* ---------------------------------------------------------------------- */
/* Manual entry modal                                                     */
/* ---------------------------------------------------------------------- */

function AccessGateModal({ onClose, onSubmit }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState(false);

  const attempt = () => {
    const ok = onSubmit(username, password);
    if (!ok) {
      setError(true);
      setPassword("");
    }
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "#000000cc", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200, padding: 20 }} onClick={onClose}>
      <div style={{ background: PALETTE.panel, border: `1px solid ${PALETTE.border}`, borderRadius: 14, width: "100%", maxWidth: 360, display: "flex", flexDirection: "column", animation: "fadeIn 0.15s ease-out" }} onClick={(e) => e.stopPropagation()}>
        <div style={{ padding: "18px 22px", borderBottom: `1px solid ${PALETTE.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, fontWeight: 700, fontSize: 15 }}>
            <Lock size={16} color={PALETTE.amber} /> Admin sign-in required
          </div>
          <button onClick={onClose} style={{ background: "none", border: "none", color: PALETTE.textDim, cursor: "pointer" }}><X size={18} /></button>
        </div>
        <div style={{ padding: "18px 22px", display: "flex", flexDirection: "column", gap: 12 }}>
          <div style={{ fontSize: 11.5, color: PALETTE.textDim, lineHeight: 1.5 }}>
            Logging months, uploading files, and exporting data are restricted. Every attempt — successful or not — is recorded.
          </div>
          <div>
            <label style={labelStyle}>Username</label>
            <input type="email" value={username} onChange={(e) => { setUsername(e.target.value); setError(false); }} placeholder="you@example.com" style={inputStyle} autoFocus />
          </div>
          <div>
            <label style={labelStyle}>Password</label>
            <input
              type="password"
              inputMode="numeric"
              maxLength={4}
              value={password}
              onChange={(e) => { setPassword(e.target.value); setError(false); }}
              onKeyDown={(e) => { if (e.key === "Enter") attempt(); }}
              placeholder="4-digit code"
              style={inputStyle}
            />
          </div>
          {error && <div style={{ fontSize: 11.5, color: PALETTE.red }}>Incorrect username or password. This attempt has been logged.</div>}
        </div>
        <div style={{ padding: "16px 22px", borderTop: `1px solid ${PALETTE.border}`, display: "flex", justifyContent: "flex-end", gap: 10 }}>
          <button onClick={onClose} style={btnStyle("transparent", PALETTE.textDim, PALETTE.border)}>Cancel</button>
          <button onClick={attempt} style={btnStyle(PALETTE.amber, PALETTE.bg)}>Sign in</button>
        </div>
      </div>
    </div>
  );
}

function AccessLogModal({ onClose, entries }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "#000000cc", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200, padding: 20 }} onClick={onClose}>
      <div style={{ background: PALETTE.panel, border: `1px solid ${PALETTE.border}`, borderRadius: 14, width: "100%", maxWidth: 480, maxHeight: "80vh", display: "flex", flexDirection: "column", animation: "fadeIn 0.15s ease-out" }} onClick={(e) => e.stopPropagation()}>
        <div style={{ padding: "18px 22px", borderBottom: `1px solid ${PALETTE.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, fontWeight: 700, fontSize: 15 }}>
            <Lock size={16} color={PALETTE.textDim} /> Access log
          </div>
          <button onClick={onClose} style={{ background: "none", border: "none", color: PALETTE.textDim, cursor: "pointer" }}><X size={18} /></button>
        </div>
        <div style={{ padding: "10px 22px 18px", overflowY: "auto" }}>
          <div style={{ fontSize: 11, color: PALETTE.textDim, margin: "8px 0 14px", lineHeight: 1.5 }}>
            Every sign-in attempt to Log month / Upload / Export is recorded here — visible to anyone with this app open, so it doubles as a shared audit trail.
          </div>
          {entries.length === 0 && <div style={{ fontSize: 13, color: PALETTE.textDim, textAlign: "center", padding: 20 }}>No attempts recorded yet.</div>}
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
            <thead>
              <tr style={{ textAlign: "left", color: PALETTE.textDim, fontSize: 10.5, textTransform: "uppercase", letterSpacing: 0.4 }}>
                <th style={{ padding: "5px 6px" }}>When</th>
                <th style={{ padding: "5px 6px" }}>Username entered</th>
                <th style={{ padding: "5px 6px" }}>Result</th>
              </tr>
            </thead>
            <tbody>
              {entries.map((e, i) => (
                <tr key={i} style={{ borderTop: `1px solid ${PALETTE.border}` }}>
                  <td className="mono" style={{ padding: "6px 6px", color: PALETTE.textDim, whiteSpace: "nowrap" }}>{new Date(e.ts).toLocaleString()}</td>
                  <td style={{ padding: "6px 6px" }}>{e.username || "—"}</td>
                  <td style={{ padding: "6px 6px", fontWeight: 700, color: e.success ? "#4ADE80" : PALETTE.red }}>{e.success ? "Success" : "Failed"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function EntryModal({ onClose, onSave, existing, siteCategories, offsiteCategories, onAddVendor }) {
  const [month, setMonth] = useState(CURRENT_EDITABLE_MONTH);
  const [site, setSite] = useState(emptyCats(siteCategories));
  const [offsite, setOffsite] = useState(emptyCats(offsiteCategories));
  const [staffSite, setStaffSite] = useState(0);
  const [staffOffsite, setStaffOffsite] = useState(0);
  const [indicators, setIndicators] = useState({});
  const [openSection, setOpenSection] = useState("site");
  const [addingTo, setAddingTo] = useState(null); // "site" | "offsite" | null
  const [newVendorName, setNewVendorName] = useState("");

  const emptyIndicators = () => {
    const obj = {};
    [...LAGGING_METRICS, ...LEADING_METRICS, ...WASTE_METRICS].forEach((m) => { obj[m.key] = { site: 0, offsite: 0 }; });
    return obj;
  };

  useEffect(() => {
    const found = existing.find((m) => m.month === month);
    if (found) {
      setSite({ ...emptyCats(siteCategories), ...found.site });
      setOffsite({ ...emptyCats(offsiteCategories), ...found.offsite });
      setStaffSite(Number(found.staffPermanent?.site) || 0);
      setStaffOffsite(Number(found.staffPermanent?.offsite) || 0);
      const merged = emptyIndicators();
      Object.entries(found.indicators || {}).forEach(([k, v]) => { if (merged[k]) merged[k] = { site: Number(v.site) || 0, offsite: Number(v.offsite) || 0 }; });
      setIndicators(merged);
    } else {
      setSite(emptyCats(siteCategories));
      setOffsite(emptyCats(offsiteCategories));
      setStaffSite(0);
      setStaffOffsite(0);
      setIndicators(emptyIndicators());
    }
  }, [month, siteCategories, offsiteCategories]); // eslint-disable-line

  const siteTotal = sum(site);
  const offsiteTotal = sum(offsite);
  const staffTotal = (Number(staffSite) || 0) + (Number(staffOffsite) || 0);
  const setIndicatorField = (key, field, val) => setIndicators((v) => ({ ...v, [key]: { ...v[key], [field]: val === "" ? 0 : Number(val) } }));

  const submitNewVendor = () => {
    const name = newVendorName.trim();
    if (!name || !addingTo) return;
    onAddVendor(addingTo, name);
    if (addingTo === "site") setSite((s) => ({ ...s, [name]: 0 }));
    else setOffsite((s) => ({ ...s, [name]: 0 }));
    setNewVendorName("");
    setAddingTo(null);
  };

  const indicatorGroups = [
    { id: "lagging", title: "Lagging indicators (this month)", accent: PALETTE.red, rows: LAGGING_METRICS },
    { id: "leading", title: "Leading indicators (this month)", accent: PALETTE.teal, rows: LEADING_METRICS },
    { id: "waste", title: "Waste disposal (this month)", accent: PALETTE.textDim, rows: WASTE_METRICS },
  ];

  return (
    <div style={{ position: "fixed", inset: 0, background: "#000000cc", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200, padding: 20 }} onClick={onClose}>
      <div style={{ background: PALETTE.panel, border: `1px solid ${PALETTE.border}`, borderRadius: 14, width: "100%", maxWidth: 640, maxHeight: "88vh", display: "flex", flexDirection: "column", animation: "fadeIn 0.15s ease-out" }} onClick={(e) => e.stopPropagation()}>
        <div style={{ padding: "18px 22px", borderBottom: `1px solid ${PALETTE.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ fontWeight: 700, fontSize: 15 }}>Log month</div>
          <button onClick={onClose} style={{ background: "none", border: "none", color: PALETTE.textDim, cursor: "pointer" }}><X size={18} /></button>
        </div>

        <div style={{ padding: "18px 22px", overflowY: "auto" }}>
          <label style={labelStyle}>Month</label>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <input type="month" value={month} onChange={(e) => setMonth(e.target.value)} style={{ ...inputStyle, flex: 1 }} />
            {month === CURRENT_EDITABLE_MONTH ? (
              <span style={{ fontSize: 11, fontWeight: 700, color: PALETTE.amber, background: `${PALETTE.amber}1a`, border: `1px solid ${PALETTE.amberDim}`, padding: "6px 10px", borderRadius: 7, whiteSpace: "nowrap" }}>Current month</span>
            ) : (
              <button onClick={() => setMonth(CURRENT_EDITABLE_MONTH)} style={{ fontSize: 11, fontWeight: 600, color: PALETTE.textDim, background: "none", border: `1px solid ${PALETTE.border}`, padding: "6px 10px", borderRadius: 7, cursor: "pointer", whiteSpace: "nowrap" }}>Jump to current</button>
            )}
          </div>
          <div style={{ fontSize: 11.5, color: PALETTE.textDim, marginTop: 8, lineHeight: 1.5 }}>
            Everything you enter here adds to the running PTD and YTD totals shown across the dashboard — nothing is typed in as a fixed total.
          </div>

          <SectionToggle title="EcoOils Staff Permanent" colorAccent={PALETTE.purple} total={staffTotal} open={openSection === "staff"} onToggle={() => setOpenSection(openSection === "staff" ? "" : "staff")}>
            <div style={gridStyle}>
              <NumField label="Site" value={staffSite} onChange={setStaffSite} />
              <NumField label="Offsite / Office" value={staffOffsite} onChange={setStaffOffsite} />
            </div>
          </SectionToggle>

          <SectionToggle title="Site hours — Sei Mangkei (contractors)" colorAccent={SITE_COLOR} total={siteTotal} open={openSection === "site"} onToggle={() => setOpenSection(openSection === "site" ? "" : "site")}>
            <div style={gridStyle}>
              {siteCategories.map((c) => (
                <NumField key={c} label={vendorLabel(c) !== c.trim() ? c.trim() : c} value={site[c]} onChange={(v) => setSite((s) => ({ ...s, [c]: v }))} />
              ))}
            </div>
            <AddVendorRow active={addingTo === "site"} value={newVendorName} onChange={setNewVendorName} onStart={() => setAddingTo("site")} onCancel={() => { setAddingTo(null); setNewVendorName(""); }} onSubmit={submitNewVendor} />
          </SectionToggle>

          <SectionToggle title="Offsite / Office hours (contractors)" colorAccent={OFFSITE_COLOR} total={offsiteTotal} open={openSection === "offsite"} onToggle={() => setOpenSection(openSection === "offsite" ? "" : "offsite")}>
            <div style={gridStyle}>
              {offsiteCategories.map((c) => (
                <NumField key={c} label={vendorLabel(c) !== c.trim() ? c.trim() : c} value={offsite[c]} onChange={(v) => setOffsite((s) => ({ ...s, [c]: v }))} />
              ))}
            </div>
            <AddVendorRow active={addingTo === "offsite"} value={newVendorName} onChange={setNewVendorName} onStart={() => setAddingTo("offsite")} onCancel={() => { setAddingTo(null); setNewVendorName(""); }} onSubmit={submitNewVendor} />
          </SectionToggle>

          {indicatorGroups.map((g) => (
            <SectionToggle key={g.id} title={g.title} colorAccent={g.accent} total={g.rows.length} totalLabel="metrics" open={openSection === g.id} onToggle={() => setOpenSection(openSection === g.id ? "" : g.id)}>
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {g.rows.map((m) => (
                  <div key={m.key} style={{ display: "grid", gridTemplateColumns: "1.6fr 0.7fr 0.7fr", gap: 8, alignItems: "center" }}>
                    <div style={{ fontSize: 12.5, color: PALETTE.text }}>{m.label}{m.unit ? ` (${m.unit})` : ""}</div>
                    <NumField label="Site" value={indicators[m.key]?.site ?? 0} onChange={(v) => setIndicatorField(m.key, "site", v)} />
                    <NumField label="Offsite" value={indicators[m.key]?.offsite ?? 0} onChange={(v) => setIndicatorField(m.key, "offsite", v)} />
                  </div>
                ))}
              </div>
            </SectionToggle>
          ))}
        </div>

        <div style={{ padding: "16px 22px", borderTop: `1px solid ${PALETTE.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div className="mono" style={{ fontSize: 13, color: PALETTE.textDim }}>
            Month total: <b style={{ color: PALETTE.text }}>{fmt(siteTotal + offsiteTotal + staffTotal)}</b> hrs
          </div>
          <div style={{ display: "flex", gap: 10 }}>
            <button onClick={onClose} style={btnStyle("transparent", PALETTE.textDim, PALETTE.border)}>Cancel</button>
            <button onClick={() => onSave({ month, total: siteTotal + offsiteTotal + staffTotal, site, offsite, staffPermanent: { site: Number(staffSite) || 0, offsite: Number(staffOffsite) || 0 }, indicators })} style={btnStyle(PALETTE.amber, PALETTE.bg)}>Save month</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function AddVendorRow({ active, value, onChange, onStart, onCancel, onSubmit }) {
  if (!active) {
    return (
      <button onClick={onStart} style={{ marginTop: 12, background: "none", border: `1px dashed ${PALETTE.border}`, borderRadius: 7, padding: "8px 12px", color: PALETTE.textDim, cursor: "pointer", fontSize: 12.5, display: "flex", alignItems: "center", gap: 6, width: "100%", justifyContent: "center" }}>
        <PlusCircle size={13} /> Add contractor
      </button>
    );
  }
  return (
    <div style={{ marginTop: 12, display: "flex", gap: 8 }}>
      <input
        type="text"
        autoFocus
        value={value}
        onChange={(e) => onChange(e.target.value)}
        onKeyDown={(e) => { if (e.key === "Enter") onSubmit(); if (e.key === "Escape") onCancel(); }}
        placeholder="New contractor name"
        style={{ ...inputStyle, padding: "7px 9px", fontSize: 13 }}
      />
      <button onClick={onSubmit} style={{ ...btnStyle(PALETTE.amber, PALETTE.bg), padding: "7px 12px" }}>Add</button>
      <button onClick={onCancel} style={{ ...btnStyle("transparent", PALETTE.textDim, PALETTE.border), padding: "7px 12px" }}>Cancel</button>
    </div>
  );
}

function SectionToggle({ title, colorAccent, total, totalLabel, open, onToggle, children }) {
  return (
    <div style={{ marginTop: 18, border: `1px solid ${PALETTE.border}`, borderRadius: 10, overflow: "hidden" }}>
      <button onClick={onToggle} style={{ width: "100%", display: "flex", justifyContent: "space-between", alignItems: "center", background: PALETTE.panelAlt, border: "none", padding: "12px 16px", cursor: "pointer", color: PALETTE.text }}>
        <span style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, fontWeight: 600 }}>
          <span style={{ width: 8, height: 8, borderRadius: "50%", background: colorAccent }} />
          {title}
          <span className="mono" style={{ color: PALETTE.textDim, fontWeight: 400 }}>· {fmt(total)} {totalLabel || "hrs"}</span>
        </span>
        {open ? <ChevronUp size={16} color={PALETTE.textDim} /> : <ChevronDown size={16} color={PALETTE.textDim} />}
      </button>
      {open && <div style={{ padding: 16 }}>{children}</div>}
    </div>
  );
}

function NumField({ label, value, onChange }) {
  return (
    <div>
      <div style={{ fontSize: 11, color: PALETTE.textDim, marginBottom: 4, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }} title={label}>{label}</div>
      <input
        type="number"
        value={value === 0 ? "" : value}
        placeholder="0"
        onChange={(e) => onChange(e.target.value === "" ? 0 : Number(e.target.value))}
        className="mono"
        style={{ ...inputStyle, padding: "7px 9px", fontSize: 13 }}
      />
    </div>
  );
}

const labelStyle = { display: "block", fontSize: 11, color: PALETTE.textDim, textTransform: "uppercase", letterSpacing: 0.6, marginBottom: 6 };
const inputStyle = { width: "100%", background: "#0B0F12", border: `1px solid ${PALETTE.border}`, borderRadius: 7, padding: "9px 11px", color: PALETTE.text, fontSize: 14 };
const gridStyle = { display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(120px, 1fr))", gap: 10 };
