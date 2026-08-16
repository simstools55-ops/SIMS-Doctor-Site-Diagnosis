# Weekly Aggregation Hotfix

Changed Apps Script files:
- WeeklyTrendAnalyzer.gs — REPLACE
- SiteDiagnosisConfig.gs — REPLACE

All other .gs files: NO CHANGE.
appsscript.json: NO CHANGE.

Fix:
- Aggregate page_weekly evidence by Normalized URL + Week Start.
- Sum clicks and impressions.
- Recalculate average position using impression-weighted average.
- Normalize week date keys before grouping.

Test:
1. Replace the two files.
2. Save and reload.
3. Run `4. Run Site Analysis`.
4. Run `6. Validate Weekly Trends`.
5. In `Weekly Trend Validation`, each URL should now have only one row per week.
