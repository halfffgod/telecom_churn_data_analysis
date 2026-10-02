# Telecom Customer Churn Analysis

## Project Overview

This project analyzes customer churn in the telecommunications industry using Excel. It focuses on data quality, customer churn patterns, and a financial scenario model to evaluate potential retention strategies.

The analysis is organized around three key business questions:

1. **Data Quality:** Is the dataset clean, complete, and internally consistent?
2. **Churn Analysis:** What is the overall churn rate, and how does it vary by contract type and offer?
3. **Pricing and Retention Modeling:** What discount on Monthly Charge would be required to reduce churn among Month-to-Month customers, and would the estimated financial impact justify the cost?

The project uses Excel features such as Power Query, Excel Tables, formulas, XLOOKUP, COUNTIFS, data validation checks, charts, and Goal Seek.

---

## Data Sources

The project uses two datasets:

* **Telecom Customer Churn:** Customer-level data containing demographic information, subscribed services, contract details, billing information, and churn status.
* **ZIP Code Population Reference:** A reference dataset containing population information by ZIP code, used to validate geographic coverage and support further analysis.

The raw data is preserved separately from the cleaned analytical dataset to maintain traceability and reproducibility.

### Data Preparation

Power Query was used to import the source data and establish refreshable connections. The datasets were structured as Excel Tables:

* `tblChurn` — customer churn dataset
* `tblPopulation` — ZIP-code population reference

Data-quality checks and validation formulas were maintained separately on the `QA_Checks` sheet. Cleaning rules were applied to the analytical data without modifying the original source records.

---

## Question 1: Data Quality Assessment

**Objective:** Validate the dataset's structural integrity, completeness, and internal consistency before performing the analysis.

### Data Quality Checks

| Check                          | Method                                            | Result                                                                 |
| ------------------------------ | ------------------------------------------------- | ---------------------------------------------------------------------- |
| Duplicate Customer IDs         | COUNTIF                                           | 0 duplicates                                                           |
| Blank values by column         | COUNTBLANK across 38 columns                      | Blanks identified in service, billing, and churn-related fields        |
| Blank pattern validation       | COUNTIFS using service and customer-status fields | Blanks explained by service status or churn status                     |
| Service dependency rules       | 12 COUNTIFS checks                                | 12/12 passed; 0 violations                                             |
| Negative or implausible values | COUNTIF on financial and demographic fields       | 120 negative Monthly Charge values identified; all other checks passed |
| Total Revenue integrity        | Recalculation using the data dictionary formula   | 0 mismatches                                                           |
| ZIP-code reference coverage    | XLOOKUP against population reference              | 0 unmatched ZIP codes                                                  |

### Key Findings and Cleaning Decisions

**1. Structural blanks and missing values**

Not all blank cells represented genuine missing data. Some fields were conditionally applicable depending on customer subscriptions or churn status.

Examples include:

* Multiple Lines is not applicable when Phone Service = No.
* Internet-related service fields are not applicable when Internet Service = No.
* Churn Category and Churn Reason are not applicable to customers who have not churned.

According to the data dictionary, service-dependent fields should contain documented default values rather than blanks. The following standardization was applied:

* 682 phone-service-dependent blanks
* 1,526 internet-service-dependent blanks

The 5,174 blank values in Churn Category and Churn Reason were intentionally preserved because they correspond to customers who did not churn.

**2. Negative Monthly Charge values**

The initial validation identified 120 records (approximately 1.7% of the dataset) with negative Monthly Charge values.

Since negative charges are inconsistent with the intended business meaning of the field, these records were treated as data-quality anomalies.

Instead of deleting the records or simply reversing the sign, the following imputation rule was applied:

`Implied Monthly Charge = Total Charges / Tenure in Months`

Excel formula:

```excel
=IF([@[Monthly Charge]]<0,
   [@[Total Charges]]/[@[Tenure in Months]],
   [@[Monthly Charge]])
```

This approach preserves the customer records while deriving replacement values from existing billing information.

**Limitation:** The imputed value represents the historical average monthly charge, not necessarily the customer's actual current charge. Changes in services, pricing, or subscription plans over time may affect the accuracy of this estimate. The original Monthly Charge values should therefore remain available for auditing.

### Post-Cleaning Validation

| Check                          | Before Cleaning | After Cleaning |
| ------------------------------ | --------------: | -------------: |
| Duplicate Customer IDs         |               0 |              0 |
| Service dependency violations  |               0 |              0 |
| Negative Monthly Charge values |             120 |              0 |
| Total Revenue mismatches       |               0 |              0 |
| Unmatched ZIP codes            |               0 |              0 |

The post-cleaning checks confirmed that the identified anomalies were addressed without introducing new structural or financial inconsistencies.

### Conclusion

The dataset passed the defined structural, business-rule, financial, and geographic validation checks after cleaning. Expected inapplicable values were distinguished from genuine data-quality anomalies, and the negative Monthly Charge values were handled through documented imputation.

The resulting dataset is suitable for downstream analysis, subject to the limitations of the imputation method.

---

## Question 2: Customer Churn Analysis

**Objective:** Calculate the overall churn rate and investigate how churn varies across contract types and customer offers.

### Overall Churn Results

| Metric                | Result |
| --------------------- | -----: |
| Overall Churn Rate    | 26.54% |
| New Join Rate         |  6.45% |
| Retained This Quarter | 67.02% |

### Churn by Contract Type

| Contract Type  | Churn Rate |
| -------------- | ---------: |
| Month-to-Month |     45.84% |
| One Year       |     10.71% |
| Two Year       |      2.55% |


<img width="792" height="448" alt="image" src="https://github.com/user-attachments/assets/14d52fde-c031-4dc3-8190-28b991a49000" />


Month-to-Month customers have a substantially higher observed churn rate than customers with longer-term contracts. The churn rate decreases progressively across the three contract categories.

This pattern indicates a strong association between contract type and churn. However, the analysis alone cannot establish that longer contracts directly cause lower churn, as customer characteristics, pricing, and other factors may also contribute.

### Churn by Offer

| Offer   | Churn Rate |
| ------- | ---------: |
| Offer A |      6.73% |
| Offer B |     12.26% |
| None    |     27.11% |
| Offer D |     26.74% |
| Offer E |     52.92% |


<img width="752" height="450" alt="image" src="https://github.com/user-attachments/assets/1403f71a-dcc4-4296-9729-3a3796af8a09" />


The observed churn rates vary considerably across offer categories.

* **Offers A and B:** Both have lower observed churn rates than the group with no offer.
* **Offer D:** Its churn rate is close to the no-offer baseline.
* **Offer E:** It has the highest observed churn rate and a join rate of 21.74%.

Offer E's combination of a high churn rate and high join rate may indicate that it is associated with customer acquisition rather than long-term retention. This remains a hypothesis requiring further investigation.

The lower churn rates associated with Offers A and B are also observational findings. They do not establish that these offers independently caused churn reduction. Differences in customer eligibility, offer targeting, and customer characteristics could explain part of the observed variation.

### Key Findings

* Contract type shows a pronounced association with churn.
* Month-to-Month customers represent a segment with substantially higher observed churn.
* Offer categories exhibit different churn patterns, suggesting that offer design and customer targeting warrant further analysis.
* The relationship between acquisition offers and long-term retention requires additional investigation.

These findings provide the basis for the pricing and retention scenario developed in Question 3.

---

## Question 3: Pricing and Retention Scenario for Month-to-Month Customers

**Objective:** Estimate the discount on Monthly Charge required to reduce Month-to-Month churn to a defined target and assess the potential financial impact of that scenario.

### Model Inputs

| Metric                                                |  Value |
| ----------------------------------------------------- | -----: |
| Month-to-Month Customer Count                         |  3,610 |
| Current Churn Rate                                    | 45.84% |
| Average Monthly Charge                                | $65.58 |
| Average Monthly Revenue Rate (Total Revenue / Tenure) | $89.05 |
| Target Churn Rate                                     | 28.28% |

The target churn rate is defined as the midpoint between the Month-to-Month churn rate (45.84%) and the One Year churn rate (10.71%).

### Assumption

The model uses an assumed churn sensitivity of 1.0, meaning that every 1% discount is assumed to reduce churn by 1 percentage point.

This is an illustrative assumption because the dataset does not contain direct evidence of customer responses to different discount levels.

The actual discount-response relationship would need to be measured through a controlled pilot or A/B test before implementing a real pricing strategy.

### Goal Seek Result

Using Excel Goal Seek, the required discount was estimated as:

**Required discount: 17.56%**

This is the modeled discount necessary to achieve the target churn rate under the specified sensitivity assumption.

### Estimated Financial Impact

| Metric                             |       Value |
| ---------------------------------- | ----------: |
| Monthly Discount Cost              |  $41,585.53 |
| Estimated Customers Retained       |        ≈634 |
| Estimated Monthly Revenue Retained |  $56,472.33 |
| Net Monthly Impact                 | +$14,886.80 |
| Annualized Net Impact              |   ≈$178,600 |

The monthly discount cost is calculated across the entire Month-to-Month segment. The estimated revenue retained is based on the average monthly revenue rate multiplied by the estimated number of customers retained.

### Interpretation

Under the model's assumptions, the estimated monthly revenue retained exceeds the monthly discount cost by approximately $14,887.

This produces a positive modeled net monthly impact, equivalent to approximately $178,600 annually if the monthly result were sustained for 12 months.

The model also corrects an earlier comparison issue by using a monthly revenue rate rather than cumulative historical revenue, ensuring that the estimated monthly cost and benefit are compared on the same time basis.

### Limitations

* **Assumed sensitivity:** The model does not establish a causal relationship between discounts and churn reduction. Its financial results depend heavily on the assumed sensitivity.
* **Broad discount application:** The discount is applied to the entire Month-to-Month segment, including customers who might remain regardless of the incentive.
* **Revenue approximation:** Total Revenue divided by Tenure represents an average historical monthly revenue rate, not necessarily current recurring revenue.
* **Simplified financial model:** The calculation does not fully account for changes in customer behavior, future revenue, operating costs, discount eligibility, or the timing of churn.
* **Sustained impact:** Annualized results assume that the modeled monthly impact continues for an entire year.

### Business Implications

The scenario illustrates how Excel's Goal Seek functionality can be used to connect a hypothetical retention target with pricing and financial outcomes.

A targeted retention strategy could be evaluated alongside a blanket discount. The observed churn differences across offer categories in Question 2 provide additional hypotheses for designing such a strategy.

Before implementation, the model should be calibrated with real discount-response data and tested against alternative scenarios to determine whether the estimated benefits are achievable.

---

## Tools and Excel Techniques Used

| Tool / Technique       | Application                                                          |
| ---------------------- | -------------------------------------------------------------------- |
| Power Query            | Importing source datasets and establishing refreshable connections   |
| Excel Tables           | Structuring customer and population data                             |
| COUNTIF / COUNTIFS     | Duplicate detection, blank analysis, and business-rule validation    |
| XLOOKUP                | ZIP-code reference matching                                          |
| IF                     | Conditional cleaning and Monthly Charge imputation                   |
| Excel formulas         | Revenue integrity checks, churn calculations, and financial modeling |
| Data validation checks | Assessing structural consistency and identifying anomalies           |
| Charts                 | Visualizing churn differences by contract type and offer             |
| Goal Seek              | Estimating the discount required to achieve the target churn rate    |

---

## Overall Conclusion

This project demonstrates an end-to-end Excel analytics workflow, progressing from data preparation and quality assurance to descriptive churn analysis and scenario-based financial modeling.

The analysis identified and addressed data-quality anomalies, quantified differences in churn across customer segments, and developed a pricing scenario to estimate the financial implications of a potential retention strategy.

The results highlight the importance of distinguishing observed relationships from causal effects and of documenting assumptions when using spreadsheet models to support business decisions.

The project provides a foundation for further analysis, including customer segmentation, offer effectiveness testing, targeted retention modeling, and more detailed financial scenario evaluation.
