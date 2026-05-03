/*
 * Thoughtrons Excel add-in — ribbon button handlers.
 * The full library lives in the taskpane (taskpane.html). The ribbon exposes
 * the most-used inserts plus the 4 classification banners.
 */

Office.onReady(function () {
  Office.actions.associate("insertBrandHeader",  runner("brand_header"));
  Office.actions.associate("insertKpiStrip",     runner("kpi_strip"));
  Office.actions.associate("insertRagDashboard", runner("rag_dashboard"));
  Office.actions.associate("insertRiskBlock",    runner("risk_block"));
  Office.actions.associate("insertActionBlock",  runner("action_items_block"));

  Office.actions.associate("stampUnclassified",  runner("classify_unclassified"));
  Office.actions.associate("stampInternal",      runner("classify_internal"));
  Office.actions.associate("stampRestricted",    runner("classify_restricted"));
  Office.actions.associate("stampConfidential",  runner("classify_confidential"));
});

function runner(blockKey) {
  return function (event) {
    const fn = (window.ThoughtronsBlocks || {})[blockKey];
    if (!fn) {
      console.error("Thoughtrons add-in: unknown block", blockKey);
      event.completed();
      return;
    }
    Excel.run(ctx => Promise.resolve(fn(ctx)))
      .catch(err => console.error("Thoughtrons add-in error:", err))
      .finally(() => event.completed());
  };
}
