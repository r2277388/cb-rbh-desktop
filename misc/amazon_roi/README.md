# Project Title

## An Appendage to this reporting: G:\SALES\2024 Sales Reports\Reports\MISC\Amazon ROI

## 1. Overview

This project aims to analyze the impact of AMS (Advertising Management Services) on CORE book sales. Specifically, the goal is to determine whether AMS advertising efforts lead to an increase in ordered units of books during specific periods.

## 2. Thought Process

### 2.1 Assumptions
- **AMS Sales**: Revenue generated from sales attributed to AMS advertising efforts.
- **AMS Units**: The number of units sold that can be directly attributed to AMS advertising efforts.

### 2.2 Actuals
- **AMS Spend**: The amount paid to AMS for advertising a specific ISBN during a given period.
- **Ordered Units**: The total number of units ordered for a specific ISBN during a given period, regardless of inventory availability.
- **Shipped Units**: The actual number of units shipped for a specific ISBN during a given period.

## 3. Target Variable

The target variable for this analysis is **Ordered Units**, as it represents the demand for the books, which is influenced by AMS advertising and other factors.

## 4. Data Sources

The analysis utilizes the following data sources:
- **AMS.csv**: Contains information about ISBNs and Periods with corresponding AMS Spend, flagging whether AMS advertising was active for that month.
- **Sellthru.csv**: Includes data on `Ordered Units` and `Shipped Units` for various ISBNs during specific periods.
- **Item Metadata**: Data obtained from a SQL database, providing detailed information about each ISBN, including price and publishing group.

## 5. Objective

The primary objective is to assess whether AMS advertising leads to an increase in ordered units during the months when it is active. This analysis will help determine the effectiveness of AMS in driving book sales.

## 6. Methodology

### 6.1 Data Preparation
- **Distinct ISBNs**: Only ISBNs present in the `Item Metadata` are considered in the analysis. This ensures consistency across datasets.
- **Filtering and Cleaning**: Rows with `Ordered Units` equal to zero are removed, and missing values are handled appropriately.

### 6.2 Regression Analysis
- **Model**: A linear regression model was used to analyze the relationship between `Ordered Units` (target variable) and several features, including `AMS Spend`, `Price`, `AMS Advertised`, `Title Age (Months)`, and various publishing groups.
- **Key Metrics**:
  - **Mean Squared Error (MSE)**: `135112.82`
  - **R-squared**: `0.32`

### 6.3 Key Findings
- **AMS Spend**: A positive coefficient indicates that higher AMS spending is associated with an increase in ordered units.
- **AMS Advertised**: Titles advertised by AMS showed a significant increase in ordered units, suggesting that AMS advertising is effective.
- **Price**: A negative coefficient indicates that higher prices tend to reduce the number of ordered units.
- **Publishing Groups**: Different publishing groups have varying impacts on sales, with some groups showing a strong positive association with ordered units.
- **Title Age**: Older titles tend to sell fewer units over time, as indicated by the negative coefficient for `Title Age (Months)`.

## 7. Potential Bias and Limitations

While the analysis suggests that AMS advertising has a meaningful impact on book sales, it is important to consider potential biases that could influence the results:

- **Selection Bias**: If AMS primarily advertises titles that are already expected to perform well (e.g., popular titles with historically high sales), the observed increase in sales might be due more to the titles' inherent popularity than the effectiveness of AMS advertising.
- **Endogeneity**: The decision to advertise certain titles might be influenced by factors that also affect sales (e.g., the publisher's expectations), leading to biased estimates of the AMS effect.
- **Overestimation of AMS Effect**: The regression model might attribute the success of high-volume titles to AMS advertising, even though these titles might have performed well regardless of the advertising efforts.

### Approaches to Address Bias:
- **Control for Baseline Popularity**: Including past sales performance as a control variable can help isolate the effect of AMS advertising from the titles' inherent success.
- **Propensity Score Matching (PSM)**: This technique reduces selection bias by matching treated and untreated units (e.g., AMS-advertised vs. non-AMS-advertised titles) based on observed characteristics.
- **Instrumental Variables (IV)**: Identifying a variable that influences whether a title is advertised by AMS, but does not directly affect sales, can help account for endogeneity.

## 8. Conclusion

The analysis suggests that AMS advertising has a meaningful impact on book sales, particularly in increasing the number of ordered units. While the model explains a moderate portion of the variance in ordered units (R² = 0.32), the results highlight the importance of AMS advertising in driving demand. Further analysis with additional variables or more complex models could help improve the model's explanatory power.

## 9. Future Work

- **Incorporate Additional Variables**: Explore other factors that might influence sales, such as seasonal trends, promotional activities, or competitor actions.
- **Improve Model Performance**: Consider using more advanced modeling techniques, such as ensemble methods, to increase the R² value.
- **Longitudinal Analysis**: Conduct a time-series analysis to understand the long-term impact of AMS advertising on book sales.