# Excel via Power Automate

The app talks to each Excel workbook through an HTTP-triggered Power Automate
flow (no Entra app / Graph permissions needed). Paste the flow's **HTTP POST
URL** wherever the app asks for an xlsx path (Preferences, shared cache path).

One flow per workbook table. You need three: CALM schedule, CVRI schedule,
Intro Course Log, plus one for the shared cache (`Cache` table with the
columns in `data_store.ALL_COLS`). The flow is identical for all of them; only
the file and table differ.

## 0. Prerequisites
- The workbook is in OneDrive/SharePoint and the data is an **Excel Table**
  (Insert > Table, "My table has headers"), e.g. `tblSchedule`, `tblIntro`,
  `tblCache`. Header names must match the app's columns exactly.
- On the **Cache** table set every column's number format to **Text** before
  the first write. Otherwise Excel turns `request_id` "12345" into a number
  and the upsert filter never matches.
- The "When an HTTP request is received" trigger is a **Premium** connector —
  confirm your UCSF licence includes it before building.

## 1. Trigger
**When an HTTP request is received** (Request connector)
- Who can trigger the flow: **Anyone**. Method: POST.
- Request body JSON schema:

```json
{ "type": "object", "properties": {
    "action": {"type": "string"}, "sheet": {"type": "string"},
    "key": {"type": "string"},
    "rows": {"type": "array"}, "keys": {"type": "array"} } }
```

- After saving, the **HTTP POST URL** appears here. That is what the app uses.

## 2. Switch
**Switch** (Control), On = `triggerBody()?['action']`, with four cases:
`read`, `append`, `upsert`, `delete`.

Excel steps below all use the *Excel Online (Business)* connector — pick the
same Location / Document Library / File / Table each time.

Put each case's actions inside a **Scope** (see section 7).
Feeld are nested like this 

## for #3
Switch
└─ Case: delete
   └─ Scope
      ├─ Apply to each            ← an action you add
      │   └─ Delete a row         ← dropped inside the loop
      └─ Response {"ok": true}    ← inside the Scope, below the loop (not inside it)
   └─ Response (error)            ← outside the Scope, below it, run-after "has failed"

Apply to each

In the case, click + Add an action (inside the Scope) and search Apply to each (Control).
Its one field, Select an output from previous steps, is where triggerBody()?['keys'] goes. Click the field, open the Expression tab (the fx tab), type triggerBody()?['keys'], and click OK.
Concurrency is separate. Click the ⋯ on the Apply to each header, then Settings. Turn Concurrency control on and set Degree of parallelism to 1.

Delete a row (Excel Online Business)

Click Add an action inside the loop and search Delete a row.
Choose the Location, Document Library, File and Table.
In Key Column, the dropdown may only list column names. Use the dropdown’s Enter custom value option, or the fx tab, and enter the expression triggerBody()?['key'].
In Key Value, use the fx tab and enter items('Apply_to_each'). Don’t pick the “Current item” dynamic chip if it appears; the expression is equivalent and clearer.

Response
Add it inside the Scope, below the Apply to each (not inside the loop), with status 200, header Content-Type: application/json and body {"ok": true}.

The append, upsert and read cases work the same way. Use fx for every expression, not plain text in the field.

Two different things have “key/value”-looking fields, so here is each one:

Apply to each has no key/value fields. It has one field, “Select an output from previous steps”, and that holds triggerBody()?['keys']. The concurrency setting is a toggle and a number (set to 1), not a key/value pair.

Delete a row does have Key Column and Key Value fields, and these are filled in. They sit inside the loop:

Field	What goes in it
Key Column	triggerBody()?['key'] (this resolves to request_id for the cache)
Key Value	items('Apply_to_each') (the current keys entry, i.e. the request ID to delete)

The Key Column dropdown only offers the table’s column names. If you can’t enter an expression there, either select request_id directly or use “Enter custom value”. Selecting request_id is fine for the cache, but the shared table’s key column will always be request_id anyway, so hard-coding it is okay.

## 3. Case `read`
1. **List rows present in a table** — Settings > Pagination on, threshold 5000.
2. **Response** — Status 200, header `Content-Type: application/json`,
   body `{"rows": @{body('List_rows_present_in_a_table')?['value']}}`

## 4. Case `append`
1. **Apply to each** over `triggerBody()?['rows']`.
   Settings > Concurrency control on, **Degree of parallelism = 1**
   (parallel Excel writes cause lock errors).
2. Inside: **Add a row into a table**, row = `items('Apply_to_each')`
   (use "switch to input entire object", or Peek code and set
   `"item": "@items('Apply_to_each')"`).
3. After the loop: **Response** 200, JSON, body `{"ok": true}`.

## 5. Case `upsert`
1. **Apply to each** over `triggerBody()?['rows']`, concurrency 1.
2. Inside: **List rows present in a table**
   - Filter Query:
     `concat(triggerBody()?['key'], ' eq ''', items('Apply_to_each')?[triggerBody()?['key']], '''')`
   - Top Count: 1.
3. **Condition**: `length(body('List_rows_present_in_a_table_2')?['value'])` is greater than `0`
   - Yes → **Update a row**: Key Column = `triggerBody()?['key']`,
     Key Value = `items('Apply_to_each')?[triggerBody()?['key']]`,
     Row = `items('Apply_to_each')`.
   - No → **Add a row into a table**, row = `items('Apply_to_each')`.
4. After the loop: **Response** `{"ok": true}`.

## 6. Case `delete`
1. **Apply to each** over `triggerBody()?['keys']`, concurrency 1.
2. Inside: **Delete a row**: Key Column = `triggerBody()?['key']`,
   Key Value = `items('Apply_to_each')`.
3. After the loop: **Response** `{"ok": true}`.

If a key isn't found, Delete a row errors; wrap it in its own Scope if you
want that ignored.

## 7. Error reporting
After each case's Scope add a second **Response**: Status 200,
body `{"ok": false, "error": "@{string(result('Scope'))}"}`,
Configure run after = **has failed** only. The app shows `error` to the user.

## 8. Wiring it in
Set the CALM/CVRI/Intro xlsx path (and the shared `data_file`) to the flow
URL. Test with:

```
python flow_excel.py test "<flow-url>" Cache
python flow_excel.py migrate ilab_requests_cache.csv "<flow-url>" Cache
```

`test` prints the table's headers and row count (an empty table has no
headers until a row exists).

## Notes
- The flow URL contains its secret signature — don't commit it.
- Each call is one flow run. Cache pushes batch all changed rows into one
  `upsert` call, but a large first migration will be slow.
- Excel dates come back as serial numbers; the Intro Log reader handles this.
- The app's "sheet" setting is only a label sent to the flow; the flow itself
  is bound to one table.
