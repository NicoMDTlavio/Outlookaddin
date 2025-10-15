(function () {
  const POWER_APPS_BASE =
    "https://apps.powerapps.com/play/e/default-d73a39db-6eda-495d-8000-7579f56d68b7/a/a7b19fd6-9b5a-432e-97ae-c6352bb852c3";
  const BODY_MAX = 4000;
  function setStatus(msg) {
    const el = document.getElementById("status");
    if (el) el.textContent = msg;
  }
  Office.onReady(() => {
    try {
      const item = Office.context.mailbox.item;
      if (!item) {
        setStatus("No item context available.");
        return;
      }
      const subject = (item.subject || "").toString();
      const fromName = (item.from && item.from.displayName) || "";
      const fromEmail = (item.from && item.from.emailAddress) || "";
      setStatus("Reading message body…");
      item.body.getAsync(Office.CoercionType.Text, (res) => {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          setStatus("Failed to read body: " + res.error.message);
          return;
        }
        const rawBody = (res.value || "").toString();
        const trimmedBody =
          rawBody.length > BODY_MAX ? rawBody.slice(0, BODY_MAX) : rawBody;
        const params = new URLSearchParams({
          tenantId: "d73a39db-6eda-495d-8000-7579f56d68b7",
          source: "outlook-addin-v2",
          subj: subject,
          from: fromName,
          email: fromEmail,
          body: trimmedBody,
        });
        const targetUrl = `${POWER_APPS_BASE}?${params.toString()}`;
        setStatus("Redirecting to Power Apps…");
        window.location.replace(targetUrl);
      });
    } catch (e) {
      setStatus("Initialization error: " + (e && e.message ? e.message : e));
    }
  });
})();