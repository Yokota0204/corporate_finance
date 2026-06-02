type ApiRequestOptions = {
  token?: string;
  payload?: Record<string, unknown>;
  query?: Record<string, string | number | boolean>;
  method?: HttpMethod;
  headers?: Record<string, string>;
};

class ApiRequest {
  toQueryString(params: Record<string, string | number | boolean>): string {
    return Object.entries(params)
      .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`)
      .join('&');
  }

  request<T>(url: string, options: ApiRequestOptions): T {
    try {
      const token = options.token;
      const payload = options.payload;
      const query = options.query ? `?${this.toQueryString(options.query)}` : "";

      const res = UrlFetchApp.fetch(url + query, {
        method: options.method || "get",
        headers: {
          "Content-Type": "application/json",
          "Authorization": token ? "Bearer " + token : "",
          ...options.headers,
        },
        payload: payload ? JSON.stringify(payload) : undefined,
      }).getContentText();

      return JSON.parse(res);
    } catch (e) {
      throw e;
    }
  }

  get<T>(url: string, options: ApiRequestOptions): T {
    return this.request(url, {
      ...options,
      method: "get",
    });
  }

  post<T>(url: string, options: ApiRequestOptions): T {
    return this.request(url, {
      ...options,
      method: "post",
    });
  }
}

const J_QUANTS_API_TOKEN = "api_token";

type CompanyInfo = {
  data: {
    Date: string;
    Code: string;
    CoName: string;
    CoNameEn: string;
    Sector17Code: string;
    Sector17CodeName: string;
    Sector33Code: string;
    Sector33CodeName: string;
    ScaleCategory: string;
    MarketCode: string;
    MarketCodeName: string;
  }[];
};

function getCompanyByCode(code: string) {
  try {
    return new ApiRequest().get<CompanyInfo>(
      "https://api.jquants.com/v2/equities/master",
      {
        query: { code },
        headers: { "x-api-key": J_QUANTS_API_TOKEN },
      }
    );
  } catch (e) {
    throw e;
  }
}

function getCompanyName(code: string) {
  if (!code) {
    return "";
  }
  try {
    const companyInfo = getCompanyByCode(code);
    log("log", JSON.stringify(companyInfo));
    return companyInfo.data[0].CoName;
  } catch (e) {
    throw e;
  }
}

// 財務情報を取得
type FinanceStatement = {
  data: {
    DisclosedDate: string;
    DisclosedTime: string;
    LocalCode: string;
    DisclosureNumber: string;
    TypeOfDocument: string;
    TypeOfCurrentPeriod: string;
    CurrentPeriodStartDate: string;
    CurrentPeriodEndDate: string;
    CurrentFiscalYearStartDate: string;
    CurrentFiscalYearEndDate: string;
    NextFiscalYearStartDate: string;
    NextFiscalYearEndDate: string;
    NetSales: string;
    OperatingProfit: string;
    OrdinaryProfit: string;
    Profit: string;
    EarningsPerShare: string;
    DilutedEarningsPerShare: string;
    TotalAssets: string;
    Equity: string;
    EquityToAssetRatio: string;
    BookValuePerShare: string;
    CashFlowsFromOperatingActivities: string;
    CashFlowsFromInvestingActivities: string;
    CashFlowsFromFinancingActivities: string;
    CashAndEquivalents: string;
    ResultDividendPerShare1stQuarter: string;
    ResultDividendPerShare2ndQuarter: string;
    ResultDividendPerShare3rdQuarter: string;
    ResultDividendPerShareFiscalYearEnd: string;
    ResultDividendPerShareAnnual: string;
    DistributionsPerUnit: string;
    ResultTotalDividendPaidAnnual: string;
    ResultPayoutRatioAnnual: string;
    ForecastDividendPerShare1stQuarter: string;
    ForecastDividendPerShare2ndQuarter: string;
    ForecastDividendPerShare3rdQuarter: string;
    ForecastDividendPerShareFiscalYearEnd: string;
    ForecastDividendPerShareAnnual: string;
    ForecastDistributionsPerUnit: string;
    ForecastTotalDividendPaidAnnual: string;
    ForecastPayoutRatioAnnual: string;
    NextYearForecastDividendPerShare1stQuarter: string;
    NextYearForecastDividendPerShare2ndQuarter: string;
    NextYearForecastDividendPerShare3rdQuarter: string;
    NextYearForecastDividendPerShareFiscalYearEnd: string;
    NextYearForecastDividendPerShareAnnual: string;
    NextYearForecastDistributionsPerUnit: string;
    NextYearForecastPayoutRatioAnnual: string;
    ForecastNetSales2ndQuarter: string;
    ForecastOperatingProfit2ndQuarter: string;
    ForecastOrdinaryProfit2ndQuarter: string;
    ForecastProfit2ndQuarter: string;
    ForecastEarningsPerShare2ndQuarter: string;
    NextYearForecastNetSales2ndQuarter: string;
    NextYearForecastOperatingProfit2ndQuarter: string;
    NextYearForecastOrdinaryProfit2ndQuarter: string;
    NextYearForecastProfit2ndQuarter: string;
    NextYearForecastEarningsPerShare2ndQuarter: string;
    ForecastNetSales: string;
    ForecastOperatingProfit: string;
    ForecastOrdinaryProfit: string;
    ForecastProfit: string;
    ForecastEarningsPerShare: string;
    NextYearForecastNetSales: string;
    NextYearForecastOperatingProfit: string;
    NextYearForecastOrdinaryProfit: string;
    NextYearForecastProfit: string;
    NextYearForecastEarningsPerShare: string;
    MaterialChangesInSubsidiaries: string;
    SignificantChangesInTheScopeOfConsolidation: string;
    ChangesBasedOnRevisionsOfAccountingStandard: string;
    ChangesOtherThanOnesBasedOnRevisionsOfAccountingStandard: string;
    ChangesInAccountingEstimates: string;
    RetrospectiveRestatement: string;
    NumberOfIssuedAndOutstandingSharesAtTheEndOfFiscalYearIncludingTreasuryStock: string;
    NumberOfTreasuryStockAtTheEndOfFiscalYear: string;
    AverageNumberOfShares: string;
    NonConsolidatedNetSales: string;
    NonConsolidatedOperatingProfit: string;
    NonConsolidatedOrdinaryProfit: string;
    NonConsolidatedProfit: string;
    NonConsolidatedEarningsPerShare: string;
    NonConsolidatedTotalAssets: string;
    NonConsolidatedEquity: string;
    NonConsolidatedEquityToAssetRatio: string;
    NonConsolidatedBookValuePerShare: string;
    ForecastNonConsolidatedNetSales2ndQuarter: string;
    ForecastNonConsolidatedOperatingProfit2ndQuarter: string;
    ForecastNonConsolidatedOrdinaryProfit2ndQuarter: string;
    ForecastNonConsolidatedProfit2ndQuarter: string;
    ForecastNonConsolidatedEarningsPerShare2ndQuarter: string;
    NextYearForecastNonConsolidatedNetSales2ndQuarter: string;
    NextYearForecastNonConsolidatedOperatingProfit2ndQuarter: string;
    NextYearForecastNonConsolidatedOrdinaryProfit2ndQuarter: string;
    NextYearForecastNonConsolidatedProfit2ndQuarter: string;
    NextYearForecastNonConsolidatedEarningsPerShare2ndQuarter: string;
    ForecastNonConsolidatedNetSales: string;
    ForecastNonConsolidatedOperatingProfit: string;
    ForecastNonConsolidatedOrdinaryProfit: string;
    ForecastNonConsolidatedProfit: string;
    ForecastNonConsolidatedEarningsPerShare: string;
    NextYearForecastNonConsolidatedNetSales: string;
    NextYearForecastNonConsolidatedOperatingProfit: string;
    NextYearForecastNonConsolidatedOrdinaryProfit: string;
    NextYearForecastNonConsolidatedProfit: string;
    NextYearForecastNonConsolidatedEarningsPerShare: string;
  }[];
  pagination_key: string;
};

function getFinanceStatement(code: string) {
  try {
    return new ApiRequest().get<FinanceStatement>(
      "https://api.jquants.com/v2/fins/summary",
      {
        query: { code },
        headers: { "x-api-key": J_QUANTS_API_TOKEN },
      }
    );
  } catch (e) {
    throw e;
  }
}
