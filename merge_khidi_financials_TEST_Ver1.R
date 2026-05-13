# ============================================================
# KHIDI 병원 재무제표 세로형 파일 -> 가로형 데이터셋 자동 병합 스크립트
# 대상 파일 예시:
#   B1_KHIDI_2020_서울대학교병원_재무상태표.xls
#   B1_KHIDI_2020_서울대학교병원_손익계산서.xls
# 병원번호 매핑 파일 예시:
#   D1_HIRA_국립대_상급종합병원(12개)_현황리스트.xls
# ============================================================

options(encoding = "UTF-8")

# 1) 필요한 패키지 설치 및 로드 -------------------------------------------
required_pkgs <- c(
  "readxl", "dplyr", "tidyr", "stringr", "readr", "purrr", "writexl", "haven"
)

missing_pkgs <- required_pkgs[!vapply(required_pkgs, requireNamespace, logical(1), quietly = TRUE)]
if (length(missing_pkgs) > 0) {
  install.packages(missing_pkgs, repos = "https://cloud.r-project.org")
}

invisible(lapply(required_pkgs, library, character.only = TRUE))

# 2) 사용자가 수정할 경로 ---------------------------------------------------
# 모든 B1_KHIDI_*.xls 파일 120개와 D1_HIRA_*.xls 파일을 같은 폴더에 넣은 뒤,
# 아래 data_dir만 본인 PC 경로로 바꾸면 됩니다.
data_dir <- "C:\\Users\\HALLYM\\Downloads\\병원경영분석\\1_기말프로젝트\\0_자료모음\\B_KHIDI_재무상태표&손익계산서"   # 예: "C:/Users/사용자/Desktop/KHIDI"
out_dir  <- file.path(data_dir, "merged_output")

# 분석 대상 연도. 필요 시 수정 가능합니다.
years_expected <- 2020:2024

# TRUE로 바꾸면 12개 병원 x 5년 x 2종 = 120개 파일이 모두 없을 때 중단합니다.
# 처음 테스트할 때는 FALSE 권장.
strict_file_check <- FALSE

dir.create(out_dir, showWarnings = FALSE, recursive = TRUE)

# 3) 보조 함수 --------------------------------------------------------------
normalize_key <- function(x) {
  x |>
    as.character() |>
    stringr::str_squish() |>
    stringr::str_replace_all("\\s+", "")
}

clean_account <- function(x) {
  x |>
    as.character() |>
    stringr::str_replace_all("\u00A0", " ") |>
    stringr::str_squish()
}

parse_amount <- function(x) {
  # 원 단위 금액을 숫자로 변환합니다. 쉼표, 공백, 원 표시가 있어도 처리합니다.
  x_chr <- as.character(x)
  x_chr <- stringr::str_replace_all(x_chr, "[,\\s]", "")
  x_chr <- dplyr::na_if(x_chr, "")
  x_chr <- dplyr::na_if(x_chr, "-")
  suppressWarnings(readr::parse_number(x_chr, locale = readr::locale(grouping_mark = ",")))
}

read_hospital_map <- function(path) {
  raw <- readxl::read_excel(path, col_names = FALSE, col_types = "text", trim_ws = TRUE)
  if (ncol(raw) < 2) {
    stop("병원 매핑 파일은 최소 2개 열, 즉 병원명과 병원번호가 필요합니다: ", path)
  }

  raw <- raw[, 1:2]
  names(raw) <- c("hospital_name", "hospital_id")

  out <- raw |>
    dplyr::mutate(
      hospital_name = stringr::str_squish(as.character(hospital_name)),
      hospital_id   = readr::parse_integer(as.character(hospital_id)),
      hospital_key  = normalize_key(hospital_name)
    ) |>
    dplyr::filter(!is.na(hospital_name), hospital_name != "", !is.na(hospital_id)) |>
    dplyr::distinct(hospital_key, .keep_all = TRUE) |>
    dplyr::select(hospital_id, hospital_name, hospital_key) |>
    dplyr::arrange(hospital_id)

  if (anyDuplicated(out$hospital_id) > 0) {
    stop("D1 병원 매핑 파일에 중복 병원번호가 있습니다. 먼저 D1 파일을 확인하세요.")
  }

  out
}

get_file_meta <- function(path) {
  fname <- basename(path)

  year <- stringr::str_extract(fname, "(?<=KHIDI_)20\\d{2}")

  stmt_name <- dplyr::case_when(
    stringr::str_detect(fname, "재무상태표") ~ "재무상태표",
    stringr::str_detect(fname, "손익계산서") ~ "손익계산서",
    TRUE ~ NA_character_
  )

  stmt_type <- dplyr::case_when(
    stmt_name == "재무상태표" ~ "BS",
    stmt_name == "손익계산서" ~ "IS",
    TRUE ~ NA_character_
  )

  hospital_file_name <- fname |>
    stringr::str_remove("^(.*_)?KHIDI_20\\d{2}_") |>
    stringr::str_remove(stringr::regex("_(재무상태표|손익계산서)\\.xls[x]?$", ignore_case = TRUE)) |>
    stringr::str_squish()

  tibble::tibble(
    file_path = path,
    file_name = fname,
    year = as.integer(year),
    hospital_file_name = hospital_file_name,
    hospital_key = normalize_key(hospital_file_name),
    stmt_type = stmt_type,
    stmt_name = stmt_name
  )
}

read_statement_one <- function(path) {
  meta <- get_file_meta(path)

  if (is.na(meta$year) || is.na(meta$stmt_type) || is.na(meta$hospital_file_name)) {
    stop("파일명에서 연도/병원명/재무제표 종류를 추출하지 못했습니다: ", basename(path))
  }

  raw <- readxl::read_excel(path, col_names = FALSE, col_types = "text", trim_ws = FALSE)
  raw <- tibble::as_tibble(raw)

  if (nrow(raw) == 0 || ncol(raw) == 0) {
    stop("빈 파일입니다: ", basename(path))
  }

  # 각 셀을 문자형으로 정리한 뒤, '계정과목'이 있는 행과 열을 찾습니다.
  raw_cell <- raw |>
    dplyr::mutate(dplyr::across(dplyr::everything(), ~ stringr::str_squish(as.character(.x))))

  header_rows <- which(apply(raw_cell, 1, function(r) any(r == "계정과목", na.rm = TRUE)))
  if (length(header_rows) == 0) {
    stop("'계정과목' 행을 찾지 못했습니다: ", basename(path))
  }

  header_row <- header_rows[1]
  header_vals <- as.character(unlist(raw_cell[header_row, ], use.names = FALSE))

  account_col <- which(header_vals == "계정과목")[1]
  current_candidates <- which(stringr::str_detect(header_vals, "당\\)?기|당기"))
  current_col <- if (length(current_candidates) > 0) current_candidates[1] else account_col + 1L

  if (is.na(account_col) || is.na(current_col) || current_col > ncol(raw)) {
    stop("계정과목 열 또는 제(당)기 금액 열을 찾지 못했습니다: ", basename(path))
  }

  dat <- raw[(header_row + 1):nrow(raw), c(account_col, current_col), drop = FALSE]
  names(dat) <- c("account_raw", "amount_raw")

  dat |>
    dplyr::mutate(
      account_name = clean_account(account_raw),
      amount = parse_amount(amount_raw)
    ) |>
    dplyr::filter(!is.na(account_name), account_name != "", !is.na(amount)) |>
    dplyr::mutate(item_seq = dplyr::row_number()) |>
    dplyr::select(item_seq, account_name, amount) |>
    dplyr::mutate(
      file_path = meta$file_path,
      file_name = meta$file_name,
      year = meta$year,
      hospital_file_name = meta$hospital_file_name,
      hospital_key = meta$hospital_key,
      stmt_type = meta$stmt_type,
      stmt_name = meta$stmt_name,
      .before = 1
    )
}

make_wide_dataset <- function(long_all, stmt_type_value, prefix) {
  long_stmt <- long_all |>
    dplyr::filter(stmt_type == stmt_type_value)

  if (nrow(long_stmt) == 0) {
    stop("해당 재무제표 파일이 없습니다: ", stmt_type_value)
  }

  # 재무상태표에는 '(국고보조금)'처럼 동일 계정명이 여러 번 반복됩니다.
  # 그래서 계정명을 변수명으로 직접 쓰지 않고, 항목 순서 기반 변수명(bs_001 등)을 만들고
  # codebook에 실제 계정명을 저장합니다.
  codebook <- long_stmt |>
    dplyr::group_by(item_seq) |>
    dplyr::summarise(
      account_label_first = dplyr::first(account_name),
      account_labels_all = paste(unique(account_name), collapse = " | "),
      n_labels = dplyr::n_distinct(account_name),
      n_files = dplyr::n_distinct(file_path),
      .groups = "drop"
    ) |>
    dplyr::arrange(item_seq) |>
    dplyr::mutate(
      var_name = sprintf("%s_%03d", prefix, item_seq),
      label_check = dplyr::if_else(n_labels == 1, "OK", "CHECK_ITEM_LABEL_MISMATCH")
    ) |>
    dplyr::select(var_name, item_seq, account_label_first, account_labels_all, n_labels, n_files, label_check)

  long_stmt2 <- long_stmt |>
    dplyr::left_join(codebook |> dplyr::select(item_seq, var_name), by = "item_seq")

  duplicated_items <- long_stmt2 |>
    dplyr::count(hospital_id, hospital_name, year, stmt_type, var_name) |>
    dplyr::filter(n > 1)

  if (nrow(duplicated_items) > 0) {
    readr::write_excel_csv(duplicated_items, file.path(out_dir, paste0(prefix, "_duplicated_items.csv")))
    stop("동일 병원-연도-항목이 중복되었습니다. out_dir의 duplicated_items 파일을 확인하세요.")
  }

  wide <- long_stmt2 |>
    dplyr::select(hospital_id, hospital_name, year, var_name, amount) |>
    tidyr::pivot_wider(names_from = var_name, values_from = amount) |>
    dplyr::arrange(hospital_id, year)

  list(
    wide = wide,
    codebook = codebook,
    mismatch = codebook |> dplyr::filter(label_check != "OK")
  )
}

add_spss_labels <- function(df, codebook, statement_label) {
  attr(df$hospital_id, "label") <- "D1_HIRA 현황리스트의 병원번호"
  attr(df$hospital_name, "label") <- "병원명"
  attr(df$year, "label") <- "회계연도"

  for (i in seq_len(nrow(codebook))) {
    v <- codebook$var_name[i]
    if (v %in% names(df)) {
      attr(df[[v]], "label") <- substr(
        paste0(statement_label, " | ", codebook$account_label_first[i]),
        1,
        120
      )
    }
  }
  df
}

# 4) 파일 탐색 및 병원번호 매핑 -------------------------------------------
map_candidates <- list.files(
  data_dir,
  pattern = "^D1_HIRA_.*\\.xls[x]?$",
  full.names = TRUE,
  ignore.case = TRUE
)
if (length(map_candidates) == 0) {
  stop("D1_HIRA 병원 현황리스트 파일을 찾지 못했습니다. data_dir 경로를 확인하세요.")
}
map_file <- map_candidates[1]
hospital_map <- read_hospital_map(map_file)

# B1만 찾지 않고 B1~B12 전체를 찾도록 넓게 탐색합니다.
statement_files <- list.files(
  data_dir,
  pattern = "\\.xls[x]?$",
  full.names = TRUE,
  ignore.case = TRUE
)

statement_files <- statement_files[!grepl("^~\\$", basename(statement_files))]
statement_files <- statement_files[grepl("KHIDI", basename(statement_files), ignore.case = TRUE)]
statement_files <- statement_files[grepl("20[0-9]{2}", basename(statement_files))]
statement_files <- statement_files[grepl("재무상태표|손익계산서", basename(statement_files))]

readr::write_excel_csv(
  tibble::tibble(
    file_name = basename(statement_files),
    file_path = statement_files
  ),
  file.path(out_dir, "validation_detected_statement_files.csv")
)

message("인식된 KHIDI 재무제표 파일 수: ", length(statement_files))

if (length(statement_files) == 0) {
  stop("KHIDI 재무제표 파일을 찾지 못했습니다. data_dir 경로와 파일명을 확인하세요.")
}

file_meta <- purrr::map_dfr(statement_files, get_file_meta) |>
  dplyr::left_join(
    hospital_map |> dplyr::select(hospital_id, hospital_name, hospital_key),
    by = "hospital_key"
  ) |>
  dplyr::arrange(stmt_type, hospital_id, year)

unmapped <- file_meta |>
  dplyr::filter(is.na(hospital_id)) |>
  dplyr::select(file_name, hospital_file_name)
if (nrow(unmapped) > 0) {
  readr::write_excel_csv(unmapped, file.path(out_dir, "validation_unmapped_hospitals.csv"))
  stop("D1 매핑 파일에서 병원번호를 찾지 못한 파일이 있습니다. validation_unmapped_hospitals.csv를 확인하세요.")
}

duplicated_files <- file_meta |>
  dplyr::count(stmt_type, hospital_id, hospital_name, year) |>
  dplyr::filter(n > 1)
if (nrow(duplicated_files) > 0) {
  readr::write_excel_csv(duplicated_files, file.path(out_dir, "validation_duplicated_files.csv"))
  stop("동일 병원-연도-재무제표 종류 파일이 중복되었습니다. validation_duplicated_files.csv를 확인하세요.")
}

expected_files <- tidyr::expand_grid(
  hospital_map |> dplyr::select(hospital_id, hospital_name),
  year = years_expected,
  stmt_type = c("BS", "IS")
)

missing_files <- expected_files |>
  dplyr::anti_join(
    file_meta |> dplyr::select(hospital_id, year, stmt_type),
    by = c("hospital_id", "year", "stmt_type")
  ) |>
  dplyr::mutate(stmt_name = dplyr::if_else(stmt_type == "BS", "재무상태표", "손익계산서")) |>
  dplyr::arrange(stmt_type, hospital_id, year)

readr::write_excel_csv(file_meta, file.path(out_dir, "validation_file_list.csv"))
readr::write_excel_csv(missing_files, file.path(out_dir, "validation_missing_files.csv"))

if (strict_file_check && nrow(missing_files) > 0) {
  stop("필수 파일이 누락되었습니다. validation_missing_files.csv를 확인하세요.")
}

# 5) 모든 재무제표 읽기 -----------------------------------------------------
long_all <- purrr::map_dfr(statement_files, read_statement_one) |>
  dplyr::left_join(
    hospital_map |> dplyr::select(hospital_id, hospital_name, hospital_key),
    by = "hospital_key"
  ) |>
  dplyr::select(
    hospital_id, hospital_name, year, stmt_type, stmt_name,
    item_seq, account_name, amount,
    file_name, file_path
  ) |>
  dplyr::arrange(stmt_type, hospital_id, year, item_seq)

# 6) 재무상태표 1개, 손익계산서 1개로 가로 병합 ----------------------------
bs_result <- make_wide_dataset(long_all, stmt_type_value = "BS", prefix = "bs")
is_result <- make_wide_dataset(long_all, stmt_type_value = "IS", prefix = "is")

balance_sheet_wide <- bs_result$wide
income_statement_wide <- is_result$wide
balance_sheet_codebook <- bs_result$codebook
income_statement_codebook <- is_result$codebook

item_mismatch <- dplyr::bind_rows(
  bs_result$mismatch |> dplyr::mutate(stmt_name = "재무상태표", .before = 1),
  is_result$mismatch |> dplyr::mutate(stmt_name = "손익계산서", .before = 1)
)

readr::write_excel_csv(item_mismatch, file.path(out_dir, "validation_item_label_mismatch.csv"))

# 7) 저장: CSV, XLSX, RDS, SPSS SAV ----------------------------------------
readr::write_excel_csv(balance_sheet_wide, file.path(out_dir, "KHIDI_재무상태표_wide.csv"))
readr::write_excel_csv(income_statement_wide, file.path(out_dir, "KHIDI_손익계산서_wide.csv"))
readr::write_excel_csv(balance_sheet_codebook, file.path(out_dir, "KHIDI_재무상태표_codebook.csv"))
readr::write_excel_csv(income_statement_codebook, file.path(out_dir, "KHIDI_손익계산서_codebook.csv"))
readr::write_excel_csv(long_all, file.path(out_dir, "KHIDI_long_raw_check.csv"))

saveRDS(balance_sheet_wide, file.path(out_dir, "KHIDI_재무상태표_wide.rds"))
saveRDS(income_statement_wide, file.path(out_dir, "KHIDI_손익계산서_wide.rds"))

writexl::write_xlsx(
  list(
    "재무상태표_wide" = balance_sheet_wide,
    "손익계산서_wide" = income_statement_wide,
    "재무상태표_codebook" = balance_sheet_codebook,
    "손익계산서_codebook" = income_statement_codebook,
    "파일검증" = file_meta,
    "누락파일" = missing_files,
    "항목명검증" = item_mismatch
  ),
  path = file.path(out_dir, "KHIDI_재무제표_통합_wide.xlsx")
)

haven::write_sav(
  add_spss_labels(balance_sheet_wide, balance_sheet_codebook, "재무상태표"),
  file.path(out_dir, "KHIDI_재무상태표_wide.sav")
)
haven::write_sav(
  add_spss_labels(income_statement_wide, income_statement_codebook, "손익계산서"),
  file.path(out_dir, "KHIDI_손익계산서_wide.sav")
)

# 8) 완료 메시지 ------------------------------------------------------------
message("완료되었습니다.")
message("출력 폴더: ", out_dir)
message("재무상태표 파일 수: ", sum(file_meta$stmt_type == "BS"))
message("손익계산서 파일 수: ", sum(file_meta$stmt_type == "IS"))
message("재무상태표 wide 크기: ", nrow(balance_sheet_wide), "행 x ", ncol(balance_sheet_wide), "열")
message("손익계산서 wide 크기: ", nrow(income_statement_wide), "행 x ", ncol(income_statement_wide), "열")
message("누락 파일 수: ", nrow(missing_files), "개")
message("항목명 불일치 수: ", nrow(item_mismatch), "개")
