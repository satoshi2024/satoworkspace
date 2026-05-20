import java.util.List;

/**
 * 選択中の帳票番号に一致する帳票設定情報を取得する。
 */
private ReportDetail findReportInfo(ReportForm mainForm, String selectedReportNo) {

    ReportDetail info = null;

    // 1つ目のリストから検索
    info = findReportInfoFromList(mainForm.getReportListA(), selectedReportNo);
    if (info != null) {
        return info;
    }

    // 2つ目のリストから検索
    info = findReportInfoFromList(mainForm.getReportListB(), selectedReportNo);
    if (info != null) {
        return info;
    }

    // 3つ目のリストから検索
    info = findReportInfoFromList(mainForm.getReportListC(), selectedReportNo);
    if (info != null) {
        return info;
    }

    return null;
}

/**
 * 帳票リストから帳票番号が一致する情報を取得する。
 */
private ReportDetail findReportInfoFromList(List<ReportDetail> reportList, String selectedReportNo) {

    if (reportList == null || selectedReportNo == null) {
        return null;
    }

    // 元のループ処理（型安全に修正）
    for (ReportDetail info : reportList) {
        if (info != null && isSameReportNo(selectedReportNo, info.getReportId())) {
            return info;
        }
    }

    return null;
}

/**
 * 帳票番号比較。
 * "319" と "0319" のような差異を吸収する。
 */
private boolean isSameReportNo(String no1, String no2) {

    if (no1 == null || no2 == null) {
        return false;
    }

    try {
        return Long.parseLong(no1) == Long.parseLong(no2);
    } catch (NumberFormatException e) {
        return no1.equals(no2);
    }
}