type ApiRequestOptions = {
  token?: string;
  payload?: Record<string, unknown>;
  query?: Record<string, string | number | boolean>;
  method?: HttpMethod;
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

function getRefreshToken() {
  try {
    return new ApiRequest().post<{ refreshToken: string }>(
      "https://api.jquants.com/v1/token/auth_user",
      {
        payload: {
          mailaddress: "yokota.02210301@gmail.com",
          password: "nnVu8KyFRurx",
        },
      }
    ).refreshToken;
  } catch (e) {
    throw e;
  }
}

function getIdToken(token: string): string {
  try {
    return new ApiRequest().post<{ idToken: string }>(
      "https://api.jquants.com/v1/token/auth_refresh",
      { query: { refreshtoken: token, }, }
    ).idToken;
  } catch (e) {
    throw e;
  }
}

type CompanyInfo = {
  info: {
    Date: string;
    Code: string;
    CompanyName: string;
    CompanyNameEnglish: string;
    Sector17Code: string;
    Sector17CodeName: string;
    Sector33Code: string;
    Sector33CodeName: string;
    ScaleCategory: string;
    MarketCode: string;
    MarketCodeName: string;
    MarginCode: string;
    MarginCodeName: string;
  }[];
};

function getCompanyByCode(token: string, code: string) {
  try {
    return new ApiRequest().get<CompanyInfo>(
      "https://api.jquants.com/v1/listed/info",
      { query: { code, }, token, }
    );
  } catch (e) {
    throw e;
  }
}

function getCompanyName(code: string) {
  if (!code) {
    return "";
  }
  const refToken = getRefreshToken();
  const token = getIdToken(refToken);

  try {
    const companyInfo = getCompanyByCode(token, code);
    log("log", JSON.stringify(companyInfo));
    return companyInfo.info[0].CompanyName;
  } catch (e) {
    throw e;
  }
}

// 財務情報を取得
type FinanceStatement = {
  statements: {
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
  }[],
  pagination_key: string;
};

function getFinanceStatement(token: string, code: string) {
  try {
    return new ApiRequest().get<FinanceStatement>(
      "https://api.jquants.com/v1/fins/statements",
      { query: { code, }, token, }
    );
  } catch (e) {
    throw e;
  }
}
