Office.onReady(() => {
  Office.context.mailbox.item.body.getAsync("text", (result) => {
    const subject = encodeURIComponent(Office.context.mailbox.item.subject);
    const sender = encodeURIComponent(Office.context.mailbox.item.from.displayName);
    const email = encodeURIComponent(Office.context.mailbox.item.from.emailAddress);
    const body = encodeURIComponent(result.value);

    const powerAppsUrl = `https://apps.powerapps.com/play/e/default-d73a39db-6eda-495d-8000-7579f56d68b7/a/a7b19fd6-9b5a-432e-97ae-c6352bb852c3?tenantId=d73a39db-6eda-495d-8000-7579f56d68b7&hint=12b52286-4728-482e-949a-3772bcfb3327&sourcetime=1760380932565&subject=${subject}&sender=${sender}&email=${email}&body=${body}`;

    window.location.href = powerAppsUrl;
  });
});
