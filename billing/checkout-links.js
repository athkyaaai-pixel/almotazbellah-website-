(() => {
  // Set the published Google Apps Script /exec URL after testing the deployment.
  const checkoutUrl = '';
  if (!checkoutUrl) return;

  for (const link of document.querySelectorAll('[data-billing-plan]')) {
    const url = new URL(checkoutUrl);
    url.searchParams.set('plan', link.dataset.billingPlan);
    link.href = url.href;
  }
})();
