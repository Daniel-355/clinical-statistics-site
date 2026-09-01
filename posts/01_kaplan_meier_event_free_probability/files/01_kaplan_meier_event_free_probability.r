# =============================================================================
# PILOT clinical figure: Kaplan-Meier Plot of Time to First Dermatologic Event
#
# This program is standalone. It contains the original ADTTE records, creates
# a reproducible simulated dataset with twice the original sample size, writes
# that dataset to a valid Excel workbook, performs the time-to-event analysis,
# and creates both PNG and RTF figure outputs.
# =============================================================================

options(stringsAsFactors = FALSE)

if (!requireNamespace("survival", quietly = TRUE)) {
  stop("The 'survival' package is required.")
}

if (!requireNamespace("zip", quietly = TRUE)) {
  stop("The 'zip' package is required to create the Excel workbook.")
}

# Locate the folder containing this R program.
get_script_directory <- function() {
  args <- commandArgs(trailingOnly = FALSE)
  file_arg <- grep("^--file=", args, value = TRUE)

  if (length(file_arg) == 1L) {
    return(dirname(normalizePath(sub("^--file=", "", file_arg))))
  }

  normalizePath(getwd())
}

output_directory <- get_script_directory()
stem <- "01_kaplan_meier_event_free_probability"

# -----------------------------------------------------------------------------
# Original data from the DOCX source (254 subjects)
# -----------------------------------------------------------------------------
original_csv <- "USUBJID,TRTAN,PARAM,PARAMCD,AVAL,CNSR,PARAMN\n01-701-1015,0,Time to First Dermatologic Event,TTDE,2,0,1\n01-701-1023,0,Time to First Dermatologic Event,TTDE,3,0,1\n01-701-1028,81,Time to First Dermatologic Event,TTDE,3,0,1\n01-701-1033,54,Time to First Dermatologic Event,TTDE,28,1,1\n01-701-1034,81,Time to First Dermatologic Event,TTDE,58,0,1\n01-701-1047,0,Time to First Dermatologic Event,TTDE,46,1,1\n01-701-1097,54,Time to First Dermatologic Event,TTDE,3,0,1\n01-701-1111,54,Time to First Dermatologic Event,TTDE,11,1,1\n01-701-1115,54,Time to First Dermatologic Event,TTDE,3,0,1\n01-701-1118,0,Time to First Dermatologic Event,TTDE,182,1,1\n01-701-1130,0,Time to First Dermatologic Event,TTDE,97,0,1\n01-701-1133,81,Time to First Dermatologic Event,TTDE,61,0,1\n01-701-1146,81,Time to First Dermatologic Event,TTDE,13,0,1\n01-701-1148,81,Time to First Dermatologic Event,TTDE,3,0,1\n01-701-1153,0,Time to First Dermatologic Event,TTDE,191,1,1\n01-701-1180,81,Time to First Dermatologic Event,TTDE,36,0,1\n01-701-1181,81,Time to First Dermatologic Event,TTDE,8,1,1\n01-701-1188,54,Time to First Dermatologic Event,TTDE,2,0,1\n01-701-1192,54,Time to First Dermatologic Event,TTDE,17,0,1\n01-701-1203,0,Time to First Dermatologic Event,TTDE,183,1,1\n01-701-1211,54,Time to First Dermatologic Event,TTDE,2,0,1\n01-701-1234,0,Time to First Dermatologic Event,TTDE,177,1,1\n01-701-1239,81,Time to First Dermatologic Event,TTDE,2,0,1\n01-701-1275,81,Time to First Dermatologic Event,TTDE,18,0,1\n01-701-1287,81,Time to First Dermatologic Event,TTDE,29,0,1\n01-701-1294,54,Time to First Dermatologic Event,TTDE,57,0,1\n01-701-1302,81,Time to First Dermatologic Event,TTDE,2,0,1\n01-701-1317,54,Time to First Dermatologic Event,TTDE,33,0,1\n01-701-1324,54,Time to First Dermatologic Event,TTDE,21,0,1\n01-701-1341,54,Time to First Dermatologic Event,TTDE,15,0,1\n01-701-1345,0,Time to First Dermatologic Event,TTDE,162,1,1\n01-701-1360,81,Time to First Dermatologic Event,TTDE,3,0,1\n01-701-1363,0,Time to First Dermatologic Event,TTDE,48,0,1\n01-701-1383,81,Time to First Dermatologic Event,TTDE,4,0,1\n01-701-1387,0,Time to First Dermatologic Event,TTDE,14,1,1\n01-701-1392,0,Time to First Dermatologic Event,TTDE,183,1,1\n01-701-1415,0,Time to First Dermatologic Event,TTDE,183,1,1\n01-701-1429,54,Time to First Dermatologic Event,TTDE,43,1,1\n01-701-1440,0,Time to First Dermatologic Event,TTDE,182,1,1\n01-701-1442,54,Time to First Dermatologic Event,TTDE,77,0,1\n01-701-1444,81,Time to First Dermatologic Event,TTDE,15,0,1\n01-702-1082,54,Time to First Dermatologic Event,TTDE,46,0,1\n01-703-1042,0,Time to First Dermatologic Event,TTDE,183,1,1\n01-703-1076,81,Time to First Dermatologic Event,TTDE,30,0,1\n01-703-1086,54,Time to First Dermatologic Event,TTDE,12,0,1\n01-703-1096,0,Time to First Dermatologic Event,TTDE,64,1,1\n01-703-1100,0,Time to First Dermatologic Event,TTDE,24,0,1\n01-703-1119,54,Time to First Dermatologic Event,TTDE,25,0,1\n01-703-1175,0,Time to First Dermatologic Event,TTDE,12,1,1\n01-703-1182,54,Time to First Dermatologic Event,TTDE,15,0,1\n01-703-1197,54,Time to First Dermatologic Event,TTDE,37,1,1\n01-703-1210,0,Time to First Dermatologic Event,TTDE,175,1,1\n01-703-1258,81,Time to First Dermatologic Event,TTDE,47,0,1\n01-703-1279,54,Time to First Dermatologic Event,TTDE,41,1,1\n01-703-1295,81,Time to First Dermatologic Event,TTDE,180,1,1\n01-703-1299,0,Time to First Dermatologic Event,TTDE,183,1,1\n01-703-1335,81,Time to First Dermatologic Event,TTDE,69,1,1\n01-703-1379,54,Time to First Dermatologic Event,TTDE,181,1,1\n01-703-1403,81,Time to First Dermatologic Event,TTDE,8,1,1\n01-703-1439,81,Time to First Dermatologic Event,TTDE,44,0,1\n01-704-1008,81,Time to First Dermatologic Event,TTDE,44,1,1\n01-704-1009,54,Time to First Dermatologic Event,TTDE,19,0,1\n01-704-1010,0,Time to First Dermatologic Event,TTDE,139,1,1\n01-704-1017,81,Time to First Dermatologic Event,TTDE,31,0,1\n01-704-1025,54,Time to First Dermatologic Event,TTDE,64,1,1\n01-704-1065,81,Time to First Dermatologic Event,TTDE,4,0,1\n01-704-1074,81,Time to First Dermatologic Event,TTDE,1,0,1\n01-704-1093,81,Time to First Dermatologic Event,TTDE,51,0,1\n01-704-1114,54,Time to First Dermatologic Event,TTDE,30,0,1\n01-704-1120,54,Time to First Dermatologic Event,TTDE,22,0,1\n01-704-1127,0,Time to First Dermatologic Event,TTDE,181,1,1\n01-704-1135,54,Time to First Dermatologic Event,TTDE,39,0,1\n01-704-1164,0,Time to First Dermatologic Event,TTDE,198,1,1\n01-704-1218,54,Time to First Dermatologic Event,TTDE,190,1,1\n01-704-1233,0,Time to First Dermatologic Event,TTDE,116,1,1\n01-704-1241,81,Time to First Dermatologic Event,TTDE,68,1,1\n01-704-1260,0,Time to First Dermatologic Event,TTDE,142,1,1\n01-704-1266,81,Time to First Dermatologic Event,TTDE,5,0,1\n01-704-1323,54,Time to First Dermatologic Event,TTDE,28,0,1\n01-704-1325,54,Time to First Dermatologic Event,TTDE,31,0,1\n01-704-1332,81,Time to First Dermatologic Event,TTDE,2,0,1\n01-704-1351,0,Time to First Dermatologic Event,TTDE,189,1,1\n01-704-1388,0,Time to First Dermatologic Event,TTDE,195,1,1\n01-704-1435,0,Time to First Dermatologic Event,TTDE,57,1,1\n01-704-1445,0,Time to First Dermatologic Event,TTDE,175,1,1\n01-705-1018,0,Time to First Dermatologic Event,TTDE,8,1,1\n01-705-1031,54,Time to First Dermatologic Event,TTDE,29,0,1\n01-705-1059,0,Time to First Dermatologic Event,TTDE,90,0,1\n01-705-1186,0,Time to First Dermatologic Event,TTDE,31,1,1\n01-705-1199,54,Time to First Dermatologic Event,TTDE,14,1,1\n01-705-1280,81,Time to First Dermatologic Event,TTDE,94,0,1\n01-705-1281,81,Time to First Dermatologic Event,TTDE,64,0,1\n01-705-1282,0,Time to First Dermatologic Event,TTDE,181,1,1\n01-705-1292,54,Time to First Dermatologic Event,TTDE,15,0,1\n01-705-1303,81,Time to First Dermatologic Event,TTDE,52,0,1\n01-705-1310,81,Time to First Dermatologic Event,TTDE,13,0,1\n01-705-1349,0,Time to First Dermatologic Event,TTDE,183,1,1\n01-705-1377,81,Time to First Dermatologic Event,TTDE,63,1,1\n01-705-1382,81,Time to First Dermatologic Event,TTDE,1,1,1\n01-705-1393,54,Time to First Dermatologic Event,TTDE,167,1,1\n01-705-1431,54,Time to First Dermatologic Event,TTDE,180,1,1\n01-706-1041,0,Time to First Dermatologic Event,TTDE,64,0,1\n01-706-1049,81,Time to First Dermatologic Event,TTDE,43,1,1\n01-706-1384,54,Time to First Dermatologic Event,TTDE,1,0,1\n01-707-1037,54,Time to First Dermatologic Event,TTDE,20,1,1\n01-707-1206,0,Time to First Dermatologic Event,TTDE,182,1,1\n01-708-1019,54,Time to First Dermatologic Event,TTDE,35,1,1\n01-708-1032,54,Time to First Dermatologic Event,TTDE,18,0,1\n01-708-1084,54,Time to First Dermatologic Event,TTDE,187,1,1\n01-708-1087,0,Time to First Dermatologic Event,TTDE,189,1,1\n01-708-1158,0,Time to First Dermatologic Event,TTDE,43,0,1\n01-708-1171,0,Time to First Dermatologic Event,TTDE,181,1,1\n01-708-1178,81,Time to First Dermatologic Event,TTDE,71,0,1\n01-708-1213,81,Time to First Dermatologic Event,TTDE,18,1,1\n01-708-1216,81,Time to First Dermatologic Event,TTDE,14,0,1\n01-708-1236,81,Time to First Dermatologic Event,TTDE,6,1,1\n01-708-1253,0,Time to First Dermatologic Event,TTDE,183,1,1\n01-708-1272,54,Time to First Dermatologic Event,TTDE,22,0,1\n01-708-1286,0,Time to First Dermatologic Event,TTDE,3,0,1\n01-708-1296,0,Time to First Dermatologic Event,TTDE,83,0,1\n01-708-1297,54,Time to First Dermatologic Event,TTDE,68,0,1\n01-708-1316,0,Time to First Dermatologic Event,TTDE,177,0,1\n01-708-1336,81,Time to First Dermatologic Event,TTDE,37,0,1\n01-708-1342,0,Time to First Dermatologic Event,TTDE,182,1,1\n01-708-1347,81,Time to First Dermatologic Event,TTDE,55,0,1\n01-708-1348,54,Time to First Dermatologic Event,TTDE,119,0,1\n01-708-1353,54,Time to First Dermatologic Event,TTDE,69,1,1\n01-708-1372,81,Time to First Dermatologic Event,TTDE,29,1,1\n01-708-1378,0,Time to First Dermatologic Event,TTDE,148,1,1\n01-708-1406,81,Time to First Dermatologic Event,TTDE,189,1,1\n01-708-1428,54,Time to First Dermatologic Event,TTDE,32,0,1\n01-709-1001,0,Time to First Dermatologic Event,TTDE,183,1,1\n01-709-1007,54,Time to First Dermatologic Event,TTDE,17,0,1\n01-709-1020,54,Time to First Dermatologic Event,TTDE,27,0,1\n01-709-1029,81,Time to First Dermatologic Event,TTDE,24,0,1\n01-709-1081,54,Time to First Dermatologic Event,TTDE,80,0,1\n01-709-1088,0,Time to First Dermatologic Event,TTDE,27,0,1\n01-709-1099,81,Time to First Dermatologic Event,TTDE,19,0,1\n01-709-1102,54,Time to First Dermatologic Event,TTDE,28,0,1\n01-709-1168,81,Time to First Dermatologic Event,TTDE,16,0,1\n01-709-1217,54,Time to First Dermatologic Event,TTDE,23,0,1\n01-709-1238,81,Time to First Dermatologic Event,TTDE,16,0,1\n01-709-1259,0,Time to First Dermatologic Event,TTDE,139,1,1\n01-709-1285,54,Time to First Dermatologic Event,TTDE,27,0,1\n01-709-1301,0,Time to First Dermatologic Event,TTDE,183,1,1\n01-709-1306,0,Time to First Dermatologic Event,TTDE,77,0,1\n01-709-1309,81,Time to First Dermatologic Event,TTDE,2,0,1\n01-709-1312,0,Time to First Dermatologic Event,TTDE,28,0,1\n01-709-1326,54,Time to First Dermatologic Event,TTDE,27,0,1\n01-709-1329,81,Time to First Dermatologic Event,TTDE,15,1,1\n01-709-1339,0,Time to First Dermatologic Event,TTDE,184,1,1\n01-709-1424,81,Time to First Dermatologic Event,TTDE,6,1,1\n01-710-1002,54,Time to First Dermatologic Event,TTDE,5,1,1\n01-710-1006,81,Time to First Dermatologic Event,TTDE,20,0,1\n01-710-1021,81,Time to First Dermatologic Event,TTDE,25,0,1\n01-710-1027,0,Time to First Dermatologic Event,TTDE,183,1,1\n01-710-1045,54,Time to First Dermatologic Event,TTDE,30,0,1\n01-710-1053,54,Time to First Dermatologic Event,TTDE,34,0,1\n01-710-1060,0,Time to First Dermatologic Event,TTDE,21,0,1\n01-710-1070,81,Time to First Dermatologic Event,TTDE,96,0,1\n01-710-1077,0,Time to First Dermatologic Event,TTDE,26,0,1\n01-710-1078,0,Time to First Dermatologic Event,TTDE,188,1,1\n01-710-1083,0,Time to First Dermatologic Event,TTDE,13,1,1\n01-710-1137,81,Time to First Dermatologic Event,TTDE,3,0,1\n01-710-1142,81,Time to First Dermatologic Event,TTDE,22,1,1\n01-710-1154,54,Time to First Dermatologic Event,TTDE,21,0,1\n01-710-1166,54,Time to First Dermatologic Event,TTDE,121,1,1\n01-710-1183,0,Time to First Dermatologic Event,TTDE,183,1,1\n01-710-1187,81,Time to First Dermatologic Event,TTDE,44,0,1\n01-710-1235,54,Time to First Dermatologic Event,TTDE,126,0,1\n01-710-1249,81,Time to First Dermatologic Event,TTDE,46,0,1\n01-710-1264,0,Time to First Dermatologic Event,TTDE,70,0,1\n01-710-1270,54,Time to First Dermatologic Event,TTDE,12,0,1\n01-710-1271,0,Time to First Dermatologic Event,TTDE,67,1,1\n01-710-1278,81,Time to First Dermatologic Event,TTDE,46,0,1\n01-710-1300,54,Time to First Dermatologic Event,TTDE,48,0,1\n01-710-1314,0,Time to First Dermatologic Event,TTDE,32,1,1\n01-710-1315,0,Time to First Dermatologic Event,TTDE,77,0,1\n01-710-1354,81,Time to First Dermatologic Event,TTDE,50,0,1\n01-710-1358,54,Time to First Dermatologic Event,TTDE,97,0,1\n01-710-1368,0,Time to First Dermatologic Event,TTDE,184,1,1\n01-710-1385,54,Time to First Dermatologic Event,TTDE,51,0,1\n01-710-1408,81,Time to First Dermatologic Event,TTDE,51,0,1\n01-711-1012,81,Time to First Dermatologic Event,TTDE,14,0,1\n01-711-1036,0,Time to First Dermatologic Event,TTDE,197,1,1\n01-711-1143,54,Time to First Dermatologic Event,TTDE,46,0,1\n01-711-1433,81,Time to First Dermatologic Event,TTDE,11,1,1\n01-713-1043,54,Time to First Dermatologic Event,TTDE,12,0,1\n01-713-1073,54,Time to First Dermatologic Event,TTDE,107,0,1\n01-713-1106,81,Time to First Dermatologic Event,TTDE,188,1,1\n01-713-1141,81,Time to First Dermatologic Event,TTDE,22,0,1\n01-713-1179,0,Time to First Dermatologic Event,TTDE,181,1,1\n01-713-1209,81,Time to First Dermatologic Event,TTDE,7,0,1\n01-713-1256,0,Time to First Dermatologic Event,TTDE,42,0,1\n01-713-1269,0,Time to First Dermatologic Event,TTDE,7,0,1\n01-713-1448,54,Time to First Dermatologic Event,TTDE,60,0,1\n01-714-1035,0,Time to First Dermatologic Event,TTDE,183,1,1\n01-714-1068,54,Time to First Dermatologic Event,TTDE,5,0,1\n01-714-1195,54,Time to First Dermatologic Event,TTDE,18,0,1\n01-714-1288,81,Time to First Dermatologic Event,TTDE,39,0,1\n01-714-1375,0,Time to First Dermatologic Event,TTDE,33,0,1\n01-714-1425,81,Time to First Dermatologic Event,TTDE,8,1,1\n01-715-1085,54,Time to First Dermatologic Event,TTDE,184,1,1\n01-715-1107,54,Time to First Dermatologic Event,TTDE,71,0,1\n01-715-1155,0,Time to First Dermatologic Event,TTDE,135,1,1\n01-715-1207,0,Time to First Dermatologic Event,TTDE,196,1,1\n01-715-1319,81,Time to First Dermatologic Event,TTDE,30,1,1\n01-715-1321,81,Time to First Dermatologic Event,TTDE,70,0,1\n01-715-1397,0,Time to First Dermatologic Event,TTDE,184,1,1\n01-715-1405,54,Time to First Dermatologic Event,TTDE,8,1,1\n01-716-1024,0,Time to First Dermatologic Event,TTDE,196,1,1\n01-716-1026,0,Time to First Dermatologic Event,TTDE,16,0,1\n01-716-1030,81,Time to First Dermatologic Event,TTDE,14,1,1\n01-716-1044,0,Time to First Dermatologic Event,TTDE,191,1,1\n01-716-1063,54,Time to First Dermatologic Event,TTDE,57,0,1\n01-716-1071,81,Time to First Dermatologic Event,TTDE,55,1,1\n01-716-1094,54,Time to First Dermatologic Event,TTDE,39,1,1\n01-716-1103,54,Time to First Dermatologic Event,TTDE,112,0,1\n01-716-1108,0,Time to First Dermatologic Event,TTDE,79,0,1\n01-716-1151,54,Time to First Dermatologic Event,TTDE,43,0,1\n01-716-1157,54,Time to First Dermatologic Event,TTDE,24,0,1\n01-716-1160,0,Time to First Dermatologic Event,TTDE,35,0,1\n01-716-1167,54,Time to First Dermatologic Event,TTDE,41,0,1\n01-716-1177,0,Time to First Dermatologic Event,TTDE,185,1,1\n01-716-1189,81,Time to First Dermatologic Event,TTDE,89,0,1\n01-716-1229,81,Time to First Dermatologic Event,TTDE,1,0,1\n01-716-1298,54,Time to First Dermatologic Event,TTDE,15,0,1\n01-716-1308,0,Time to First Dermatologic Event,TTDE,24,0,1\n01-716-1311,54,Time to First Dermatologic Event,TTDE,15,0,1\n01-716-1364,81,Time to First Dermatologic Event,TTDE,46,0,1\n01-716-1373,81,Time to First Dermatologic Event,TTDE,2,0,1\n01-716-1418,81,Time to First Dermatologic Event,TTDE,1,0,1\n01-716-1441,0,Time to First Dermatologic Event,TTDE,110,0,1\n01-716-1447,81,Time to First Dermatologic Event,TTDE,30,0,1\n01-717-1004,54,Time to First Dermatologic Event,TTDE,2,0,1\n01-717-1109,81,Time to First Dermatologic Event,TTDE,20,0,1\n01-717-1174,81,Time to First Dermatologic Event,TTDE,5,0,1\n01-717-1201,0,Time to First Dermatologic Event,TTDE,1,0,1\n01-717-1344,0,Time to First Dermatologic Event,TTDE,63,1,1\n01-717-1357,81,Time to First Dermatologic Event,TTDE,23,0,1\n01-717-1446,54,Time to First Dermatologic Event,TTDE,19,0,1\n01-718-1066,54,Time to First Dermatologic Event,TTDE,24,1,1\n01-718-1079,54,Time to First Dermatologic Event,TTDE,17,0,1\n01-718-1101,81,Time to First Dermatologic Event,TTDE,167,1,1\n01-718-1139,0,Time to First Dermatologic Event,TTDE,183,1,1\n01-718-1150,0,Time to First Dermatologic Event,TTDE,56,0,1\n01-718-1170,54,Time to First Dermatologic Event,TTDE,49,1,1\n01-718-1172,0,Time to First Dermatologic Event,TTDE,70,1,1\n01-718-1250,54,Time to First Dermatologic Event,TTDE,100,0,1\n01-718-1254,54,Time to First Dermatologic Event,TTDE,35,0,1\n01-718-1328,81,Time to First Dermatologic Event,TTDE,51,0,1\n01-718-1355,0,Time to First Dermatologic Event,TTDE,9,0,1\n01-718-1371,81,Time to First Dermatologic Event,TTDE,29,0,1\n01-718-1427,81,Time to First Dermatologic Event,TTDE,42,0,1"

original_adtte <- read.csv(
  text = original_csv,
  check.names = FALSE,
  colClasses = c(
    USUBJID = "character",
    TRTAN = "integer",
    PARAM = "character",
    PARAMCD = "character",
    AVAL = "numeric",
    CNSR = "integer",
    PARAMN = "integer"
  )
)

stopifnot(nrow(original_adtte) == 254L)

# -----------------------------------------------------------------------------
# Simulate twice the original sample size while preserving the source logic
# -----------------------------------------------------------------------------
set.seed(2026082801)

source_index <- sample(
  seq_len(nrow(original_adtte)),
  size = 2L * nrow(original_adtte),
  replace = TRUE
)

simulated_adtte <- original_adtte[source_index, , drop = FALSE]
simulated_adtte$USUBJID <- sprintf("PILOT-%04d", seq_len(nrow(simulated_adtte)))

# Apply a modest log-normal perturbation to the observed/censored time.
simulated_adtte$AVAL <- pmax(
  1,
  round(simulated_adtte$AVAL * exp(rnorm(nrow(simulated_adtte), 0, 0.08)))
)

# Retain the original treatment, endpoint, and censoring distributions.
simulated_adtte$TRTAN <- as.integer(simulated_adtte$TRTAN)
simulated_adtte$CNSR <- as.integer(simulated_adtte$CNSR)
simulated_adtte$PARAMN <- as.integer(simulated_adtte$PARAMN)
names(simulated_adtte) <- toupper(names(simulated_adtte))

stopifnot(
  nrow(simulated_adtte) == 2L * nrow(original_adtte),
  all(names(simulated_adtte) == toupper(names(simulated_adtte))),
  all(simulated_adtte$CNSR %in% c(0L, 1L))
)

simulated_adtte$TRT <- factor(
  simulated_adtte$TRTAN,
  levels = c(0, 54, 81),
  labels = c("Placebo", "PILOT 54 mg", "PILOT 81 mg")
)

# -----------------------------------------------------------------------------
# Write a valid one-sheet XLSX workbook without LibreOffice or Java
# -----------------------------------------------------------------------------
xml_escape <- function(x) {
  x <- gsub("&", "&amp;", as.character(x), fixed = TRUE)
  x <- gsub("<", "&lt;", x, fixed = TRUE)
  x <- gsub(">", "&gt;", x, fixed = TRUE)
  x <- gsub('"', "&quot;", x, fixed = TRUE)
  x
}

excel_column <- function(index) {
  result <- character()
  while (index > 0L) {
    remainder <- (index - 1L) %% 26L
    result <- c(LETTERS[remainder + 1L], result)
    index <- (index - 1L) %/% 26L
  }
  paste(result, collapse = "")
}

write_xlsx <- function(data, path) {
  workbook_dir <- tempfile("pilot_xlsx_")
  dir.create(file.path(workbook_dir, "_rels"), recursive = TRUE)
  dir.create(file.path(workbook_dir, "docProps"), recursive = TRUE)
  dir.create(file.path(workbook_dir, "xl", "_rels"), recursive = TRUE)
  dir.create(file.path(workbook_dir, "xl", "worksheets"), recursive = TRUE)

  on.exit(unlink(workbook_dir, recursive = TRUE, force = TRUE), add = TRUE)

  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    '</Types>'
  ), file.path(workbook_dir, "[Content_Types].xml"), useBytes = TRUE)

  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>',
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>',
    '</Relationships>'
  ), file.path(workbook_dir, "_rels", ".rels"), useBytes = TRUE)

  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    '<sheets><sheet name="SIMULATED_ADTTE" sheetId="1" r:id="rId1"/></sheets>',
    '</workbook>'
  ), file.path(workbook_dir, "xl", "workbook.xml"), useBytes = TRUE)

  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>',
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
    '</Relationships>'
  ), file.path(workbook_dir, "xl", "_rels", "workbook.xml.rels"), useBytes = TRUE)

  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<fonts count="2"><font><sz val="10"/><name val="Arial"/></font><font><b/><sz val="10"/><name val="Arial"/></font></fonts>',
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>',
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>',
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
    '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>',
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
    '</styleSheet>'
  ), file.path(workbook_dir, "xl", "styles.xml"), useBytes = TRUE)

  cell_xml <- function(value, reference, header = FALSE) {
    style <- if (header) ' s="1"' else ""
    if (is.numeric(value) && !is.na(value)) {
      sprintf('<c r="%s"%s><v>%s</v></c>', reference, style, value)
    } else {
      sprintf('<c r="%s" t="inlineStr"%s><is><t>%s</t></is></c>', reference, style, xml_escape(value))
    }
  }

  export_data <- data[, setdiff(names(data), "TRT"), drop = FALSE]
  rows <- character(nrow(export_data) + 1L)
  header_cells <- vapply(seq_along(export_data), function(column) {
    cell_xml(names(export_data)[column], paste0(excel_column(column), "1"), TRUE)
  }, character(1))
  rows[1] <- paste0('<row r="1">', paste(header_cells, collapse = ""), '</row>')

  for (row in seq_len(nrow(export_data))) {
    cells <- vapply(seq_along(export_data), function(column) {
      cell_xml(export_data[row, column], paste0(excel_column(column), row + 1L))
    }, character(1))
    rows[row + 1L] <- paste0('<row r="', row + 1L, '">', paste(cells, collapse = ""), '</row>')
  }

  worksheet <- c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    paste0('<dimension ref="A1:', excel_column(ncol(export_data)), nrow(export_data) + 1L, '"/>'),
    '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>',
    '<sheetFormatPr defaultRowHeight="15"/>',
    '<cols><col min="1" max="1" width="16" customWidth="1"/><col min="2" max="7" width="18" customWidth="1"/></cols>',
    '<sheetData>', rows, '</sheetData>',
    paste0('<autoFilter ref="A1:', excel_column(ncol(export_data)), nrow(export_data) + 1L, '"/>'),
    '</worksheet>'
  )
  writeLines(worksheet, file.path(workbook_dir, "xl", "worksheets", "sheet1.xml"), useBytes = TRUE)

  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">',
    '<dc:title>PILOT Simulated ADTTE Data</dc:title><dc:creator>PILOT</dc:creator>',
    '</cp:coreProperties>'
  ), file.path(workbook_dir, "docProps", "core.xml"), useBytes = TRUE)

  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>R</Application></Properties>'
  ), file.path(workbook_dir, "docProps", "app.xml"), useBytes = TRUE)

  if (file.exists(path)) unlink(path)
  zip::zipr(
    zipfile = path,
    files = ".",
    root = workbook_dir,
    include_directories = FALSE
  )

  if (!file.exists(path) || file.info(path)$size == 0L) {
    stop("Unable to create Excel workbook: ", path)
  }
}

excel_path <- file.path(output_directory, paste0(stem, "_simulated_data.xlsx"))
write_xlsx(simulated_adtte, excel_path)

# -----------------------------------------------------------------------------
# Time-to-event analysis
# CNSR = 0 denotes an event; CNSR = 1 denotes censoring.
# -----------------------------------------------------------------------------
analysis_data <- simulated_adtte
analysis_data$EVENT <- 1L - analysis_data$CNSR

km_fit <- survival::survfit(
  survival::Surv(AVAL, EVENT) ~ TRT,
  data = analysis_data,
  conf.type = "log-log"
)

logrank <- survival::survdiff(
  survival::Surv(AVAL, EVENT) ~ TRT,
  data = analysis_data
)
logrank_p <- stats::pchisq(logrank$chisq, df = length(logrank$n) - 1L, lower.tail = FALSE)

treatment_colors <- c("#333333", "#0072B2", "#D55E00")
treatment_lty <- c(1, 2, 3)
treatment_labels <- levels(analysis_data$TRT)
treatment_n <- table(analysis_data$TRT)

# -----------------------------------------------------------------------------
# Create the clinical figure and number-at-risk table
# -----------------------------------------------------------------------------
png_path <- file.path(output_directory, paste0(stem, ".png"))
grDevices::png(png_path, width = 2400, height = 1500, res = 200)

layout(matrix(c(1, 2), nrow = 2), heights = c(4.4, 1.6))
par(family = "sans", mar = c(3.0, 5.0, 4.5, 2.0), xaxs = "i", yaxs = "i")

plot(
  km_fit,
  col = treatment_colors,
  lty = treatment_lty,
  lwd = 2.2,
  mark.time = TRUE,
  conf.int = FALSE,
  xlim = c(0, ceiling(max(analysis_data$AVAL) / 30) * 30),
  ylim = c(0, 1),
  xlab = "Time (Days)",
  ylab = "Event-free Probability",
  main = "Time to First Dermatologic Event",
  cex.main = 1.0,
  cex.lab = 0.9,
  cex.axis = 0.8
)
grid(nx = NA, ny = NULL, col = "#DDDDDD", lty = 3)

legend(
  "topright",
  legend = paste0(treatment_labels, " (N=", as.integer(treatment_n), ")"),
  col = treatment_colors,
  lty = treatment_lty,
  lwd = 2.2,
  bty = "n",
  cex = 0.85
)

mtext("PILOT", side = 3, line = 2.5, adj = 0, font = 2, cex = 0.9)
mtext("Safety Population", side = 3, line = 1.2, adj = 0, cex = 0.8)


max_time <- ceiling(max(analysis_data$AVAL) / 30) * 30
risk_times <- pretty(c(0, max_time), n = 7)
risk_times <- risk_times[risk_times >= 0 & risk_times <= max_time]

risk_counts <- sapply(risk_times, function(time_point) {
  vapply(levels(analysis_data$TRT), function(group) {
    sum(analysis_data$TRT == group & analysis_data$AVAL >= time_point)
  }, integer(1))
})

label_space <- max(25, 0.14 * max_time)
par(mar = c(3.7, 5.0, 0.5, 2.0), xaxs = "i")
plot(
  NA,
  xlim = c(-label_space, max_time),
  ylim = c(0.4, 3.6),
  axes = FALSE,
  xlab = "Time (Days)",
  ylab = ""
)
axis(1, at = risk_times, labels = risk_times, cex.axis = 0.8)
text(-label_space, 3.55, "Number at Risk", adj = 0, font = 2, cex = 0.8, xpd = NA)

for (group_index in seq_along(treatment_labels)) {
  y <- 4 - group_index
  text(-label_space, y, treatment_labels[group_index], adj = 0, cex = 0.75, xpd = NA)
  text(risk_times, y, labels = risk_counts[group_index, ], cex = 0.76)
}

box(bty = "l")
grDevices::dev.off()

# -----------------------------------------------------------------------------
# Embed the PNG in a landscape RTF clinical report
# -----------------------------------------------------------------------------
rtf_escape <- function(x) {
  x <- gsub("\\\\", "\\\\\\\\", x)
  x <- gsub("\\{", "\\\\{", x)
  x <- gsub("\\}", "\\\\}", x)
  x
}

write_figure_rtf <- function(image_path, rtf_path, title, figure_number) {
  image_raw <- readBin(image_path, what = "raw", n = file.info(image_path)$size)
  image_hex <- paste(sprintf("%02x", as.integer(image_raw)), collapse = "")
  image_lines <- substring(
    image_hex,
    seq(1L, nchar(image_hex), by = 120L),
    pmin(seq(1L, nchar(image_hex), by = 120L) + 119L, nchar(image_hex))
  )

  rtf <- c(
    "{\\rtf1\\ansi\\deff0",
    "{\\fonttbl{\\f0 Arial;}{\\f1 Courier New;}}",
    "\\landscape\\paperw15840\\paperh12240\\margl720\\margr720\\margt540\\margb540",
    "{\\header\\pard\\ql\\fs16 PILOT\\tab Confidential\\par}",
    "{\\footer\\pard\\ql\\fs16 Source: Simulated ADTTE data\\tab\\qr Page {\\field{\\*\\fldinst PAGE}} of {\\field{\\*\\fldinst NUMPAGES}}\\par}",
    paste0("\\pard\\qc\\b\\fs22 Figure ", figure_number, "\\par"),
    paste0("\\pard\\qc\\b\\fs22 ", rtf_escape(title), "\\par"),
    "\\pard\\qc\\b0\\fs18 Safety Population\\par\\par",
    "{\\pict\\pngblip\\picwgoal10880\\pichgoal6800",
    image_lines,
    "}",
    "\\par\\pard\\ql\\fs16 Note: Event is defined as CNSR=0; censoring is CNSR=1. Times are measured in days.\\par",
    paste0("\\pard\\ql\\fs16 Log-rank p-value: ", format.pval(logrank_p, digits = 3, eps = 0.001), ".\\par"),
    "}"
  )

  writeLines(rtf, rtf_path, useBytes = TRUE)
}

rtf_path <- file.path(output_directory, paste0(stem, ".rtf"))
write_figure_rtf(png_path, rtf_path, "Kaplan-Meier Plot of Time to First Dermatologic Event", "14.2.1")

message("Created: ", excel_path)
message("Created: ", png_path)
message("Created: ", rtf_path)
