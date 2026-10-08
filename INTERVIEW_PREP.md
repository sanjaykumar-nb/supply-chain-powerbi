# Interview prep: questions this project will draw

Answer each one out loud in under a minute, using **your** report. The notes say what a strong answer covers; they are not scripts to memorise.

## About the data and cleaning
1. **Walk me through your data cleaning.**
   Cover: UTF-8 encoding, two date formats parsed with an en-US locale, placeholder text kept in status columns before converting to null, and the shipment vs line-item split.
2. **Why two fact tables?**
   Weight and freight exist once per shipment, on its first line. One table would either double-count freight or lose line-item detail. Same dimensions, different grain.
3. **What data-quality issues did you find, and how big were they?**
   Use the numbers from your Data Quality page: freight not separately recorded, missing weights, unrecorded shipment mode, missing PO dates, and the implausible ~2-day lead time for unrecorded modes.
4. **Why "% Late vs Schedule" and not on-time delivery?**
   The data dictionary says the scheduled date isn't the client-promised date. You report what the data supports and say what extra data you'd need.

## About Power BI
5. **Difference between a calculated column and a measure?**
   A column is computed per row at refresh and stored. A measure is computed at query time in the current filter context. Your `Days vs Schedule` is a column; `% Late vs Schedule` is a measure.
6. **What does CALCULATE do in your Late Shipments measure?**
   It changes the filter context, adding the filter `Days vs Schedule > 0`, before counting rows.
7. **Why DIVIDE instead of `/`?**
   DIVIDE returns blank instead of an error when the denominator is zero or blank.
8. **Why does a Product slicer not filter the shipments visuals?**
   Product relates only to line items, and a shipment can contain several products. Bi-directional filtering could be enabled, but it risks ambiguity and wrong totals.
9. **How did you make sure your numbers were right?**
   You recomputed every headline figure independently in Python (`scripts/check_answers.py`) and matched them.
10. **Why mark a date table?**
    So time-intelligence functions work correctly and every fact filters by one consistent calendar.

## About the business
11. **Truck is cheapest per kg but late most often. What would you tell an operations manager?**
    Frame it as a trade-off between cost and reliability, depending on how urgent and how valuable the product is. Name what you'd check next.
12. **What would you do with another month on this?**
    Get client-promised dates; model late-delivery risk by vendor, mode and country; add a product-level freight allocation.
13. **How does this apply to HP's supply chain?**
    The same questions apply to any supply chain: vendor reliability, mode choice vs cost, lead-time variability, and data quality before decisions.

## Statistics questions to revise (the JD lists Statistics)
- Mean vs median: why you show both for lead time (the data is skewed by long delays).
- Distribution and outliers: what the Days vs Schedule spread looks like.
- Correlation vs causation: "RDC shipments are later" is not proof that RDC causes delays.
- How you would test whether two vendors' late rates really differ (a two-proportion test, e.g. chi-square).
