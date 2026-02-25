/* Store the locale-specific strings */

export class UIStrings {
    private static EN = {
        "h_performance_specs": "Performance Specifications",
        "lbl_import_aps": "Import APS settings from worksheet",
        "lbl_aps_abs": "APS Absolute",
        "msg_aps_abs": "Must be a number or valid Excel cell reference. e.g., D1",
        "tip_aps_abs": "Excel cell reference or APS absolute value",
        "lbl_aps_rel": "APS Relative",
        "msg_aps_rel": "Must be a number or valid Excel cell reference. e.g., D2",
        "tip_aps_rel": "Excel cell reference or APS relative value as a fraction. e.g., enter 0.1 for 10%",
        "tip_range_copy": "Copy selected range",
        "h_method_comparison" : "Comparison",
        "lbl_x_range": "X Range",
        "msg_x_range": "Must be a valid Excel range. e.g., A2:A21",
        "lbl_y_range": "Y Range",
        "msg_y_range": "Must be a valid Excel range. e.g., B2:B21",
        "lbl_output_range": "Output Range",
        "msg_output_range": "Must be a valid Excel cell reference. e.g., E1",
        "tip_output_range": "Top left cell where the regression statistics are to be saved.",
        "lbl_reg_method": "Regression Method",
        "opt_paba": "Passing-Bablock",
        "opt_dem": "Deming",
        "opt_wdem": "Weighted Deming",
        "lbl_ci_method": "Confidence Interval Method",
        "opt_default": "Default",
        "opt_bootstrap": "Bootstrap",
        "lbl_use_calc_ratio": "Use Calculated Error Ratio",
        "lbl_error_ratio": "Error Ratio",
        "lbl_output_labels": "Output Labels",
        "lbl_ba_range": "Bland-Altman Output Range",
        "msg_ba_range": "Must be a valid Excel range. e.g., C2:I15",
        "tip_ba_range": "Cell range over which the Bland-Altman chart will be displayed.",
        "lbl_diff_type": "Difference Plot Type",
        "opt_abs": "Absolute",
        "opt_rel": "Relative",
        "lbl_sc_range": "Scatter Chart Output Range",
        "msg_sc_range": "Must be a valid Excel range. e.g., C16:I19",
        "tip_sc_range": "Cell range over which the scatter chart will be displayed.",
        "lbl_chart_data_range": "Chart Data Output Range",
        "msg_chart_data_range": "Must be a valid Excel cell reference. e.g., H1",
        "tip_chart_data_range": "Top left cell where data used to construct the charts is to be saved.",
        "btn_run": "Run",
        "h_cohen_kappa": "Cohen's Kappa",
        "inf_cohen_kappa": "Specify cutoffs for each method to assess concordance. Each method must have the same number of cutoffs.",
        "lbl_concordance_range": "Concordance Output Range",
        "tip_concordance_range": "Top left cell where the concordance metrics will be saved.",
        "h_precision": "Precision",
        "h_precision_layout": "Precision Layout",
        "h_qual_comp": "Qualitative Comparison",
        "h_load_range_defaults": "Load Range Defaults",
        "inf_load_range_defaults": "This add in was developed for use with the associated Method Verification Workbook. The default cell ranges used by the workbook can be populated by clicking the button below.",
        "btn_load_defaults": "Load Defaults",
        "inf_default_template": "Alternatively, the defaults for a user defined template can defined in a worksheet and loaded from that page. The definitions need to be supplied in two columns from cells A2:B20. The first column should give the name of the attribute as it appears below. The second column must contain the values to be copied. A1 and B1 can be column headers but they are not read.",
        "opt_import_ws": "Import from worksheet",
        "lbl_layout_sheet": "Layout Sheet",
        "tip_layout_sheet": "Name of the worksheet specifying the default values.",
        "lbl_attribute": "Attribute",
        "lbl_value": "Value",
        "h_ref_int": "Reference Interval",
        "h_box_cox": "Box Cox",
        "h_about": "About",
        "inf_about": "The purpose of this add-in is to provide some statistical procedures that are necessary for assessing clinical laboratory methods but are not easy to implement using builtin Excel spreadsheet functions. It is not meant to be a comprehensive statistical analysis tool. The current interation allows the user to perform linear regression techniques including Passing-Bablok, Deming, and Weighted Deming. Procedures are also provided to allow users to analyse variance components as described in CLSI EP15 and EP05.",
        "lbl_ifu": "Instructions for use",
        "h_src": "Source",
        "inf_src": "Source code is available at",
        "inf_disclaimer": "THIS CODE IS PROVIDED AS IS WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.",
        "tab_evaluation": "Evaluation",
        "tab_utilities": "Utilities",
        "tab_help": "Help",
        "inf_precision": "Use this section to analyse the imprecision of an assay. The calculations are based on those described by CLSI EP15 and EP05.",
        "lbl_days_range": "Days Range",
        "msg_days_range": "Must be a valid Excel range. e.g., A2:A26",
        "lbl_runs_range": "Runs Range",
        "msg_runs_range": "Must be a valid Excel range. e.g., B2:B26",
        "lbl_results_range": "Results Range",
        "msg_results_range": "Must be a valid Excel range. e.g., C2:C26",
        "btn_grubb_test": "Grubb's Test For Outliers",
        
    }
    // implement at a later date
    private static JA = {
        "h_performance_specs": "Performance Specifications",
        "lbl_import_aps": "Import APS settings from worksheet",
        "lbl_aps_abs": "APS Absolute",
        "msg_aps_abs": "Must be a number or valid Excel cell reference. e.g., D1",
        "tip_aps_abs": "Excel cell reference or APS absolute value",
        "lbl_aps_rel": "APS Relative",
        "msg_aps_rel": "Must be a number or valid Excel cell reference. e.g., D2",
        "tip_aps_rel": "Excel cell reference or APS relative value as a fraction. e.g., enter 0.1 for 10%",
        "tip_range_copy": "Copy selected range",
    }
    public static getLocaleStrings = (locale: string) => {
        let text: { [key: string]: string };

        // Get the resource strings that match the language.
        switch (locale) {
            case 'en-US':
                text = UIStrings.EN;
                break;
            // case 'ja-JP':
            //     text = UIStrings.JA;
            //     break;
            default:
                text = UIStrings.EN;
                break;
        }

        return text;
    } //getLocaleStrings
} //UIStrings
