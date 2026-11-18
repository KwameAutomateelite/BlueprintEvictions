# CTO technical incident events

The published business error handler routes scoped failures to a private CTO channel, outside notice and retainer notifications. It records a durable operations receipt and wakes only the CTO's marker-based routine. It has no polling schedule and does not notify the developer.

The wake payload contains workflow/execution identity and a fixed error category. Raw exception text, source email, client details and stack traces never enter it. The CTO must inspect the actual scoped execution before diagnosis. TEST ONLY comes from the registered owned fixture identity, not arbitrary exception text.

A stable incident ID, saved notification receipt and workflow deduplication prevent an identical failure event from posting again. Storage failure remains explicit; a Slack receipt is checked before claiming notification. Business jobs are never replayed by this handler.

A controlled production-webhook failure is required to verify an Error Trigger route. A manual editor run or synthetic message alone cannot prove automatic delivery. The owned fixture must be deactivated after the test. Native CTO acknowledgment must be verified separately from a Slack post.

This incident route does not grant broader repairs. Legal/case data, credentials, client sends, signing, dispatch, paid actions and business approval gates remain outside it. Compare current state before any scoped restore; never reactivate the historical developer/notices-channel route.
