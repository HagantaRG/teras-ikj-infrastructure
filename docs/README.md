# Inventory setup and operation

## Manifest

Manage inventory items in `daftar-barang`, using these row-1 headers:

| Header | Example | Meaning |
| --- | --- | --- |
| ID Barang | TRS-0001 | Permanent, unique inventory ID |
| Nama Barang | Kopi Bubuk | Item name |
| Unit | kg | Unit used for stock counts |
| Level Stok Minim | 2 | Minimum stock; read but not yet used for alerts |
| Aktif | TRUE | Include this item in the form |

Headers are matched by name, ignoring surrounding spaces and letter case.
Column order may change. IDs must match `TRS-0000` and must never be reused
for a different item. Changing an ID creates a different inventory identity.
Names and units may change without changing the ID. An empty Aktif cell
currently means active; use FALSE or an unchecked checkbox to deactivate an item.

## Clean-start setup

1. Deploy the code to Apps Script through the repository's main-branch workflow
   or your configured clasp installation.
2. Open the spreadsheet's Apps Script editor and run `resetInventoryData()` once.
3. Run `installInventoryTriggers()` using the account that will run the automation.
4. Inspect `stok-barang`, `dashboard-data`, the existing form, and Apps Script
   Executions to confirm the setup succeeded.

The reset validates the manifest, then deletes old inventory records from row 2
onward, inventory columns C onward, dashboard rows under the existing four
headers, all questions/responses in the configured inventory form, and that
form's saved question mappings. It rebuilds the inventory and form using
manifest IDs. It preserves `daftar-barang`, A/B headings, the existing form URL,
unrelated script properties, and other worksheet tabs. Copies in an external
form-response destination are not cleared. Never install a trigger for the reset.

This is a clean-start release, not a migration. Numeric legacy mappings or
dashboard IDs are rejected. Do not run the reset again after collecting real
stock data unless you intend to erase that data.

## Routine control flow

1. A manifest edit or spreadsheet row/column insertion/deletion calls
   `createFormFromSheet()`. A manual run of that function also performs a sync.
2. Under the script lock, validate the manifest and locate columns by their
   `INVENTORY_ITEM_ID` metadata value, which contains ID Barang. Create columns
   for new active IDs and write headers as `Nama Barang/Unit`.
3. Rebuild dashboard rows for retained stock columns. Use ID Barang in the
   existing `ID Produk` field; keep the other headers `Tanggal`, `Produk`, and
   `Jumlah`. Inactive item history remains available for corrections.
4. Update the existing form. Keep question IDs for renamed active items and
   remove questions for inactive or removed items. All stock questions remain
   optional and accept whole numbers of zero or greater.
5. On submission, resolve question ID to ID Barang to the current stock column.
   Write nonblank answers to today's date in column B, creating a row if needed.
   Blank answers leave cells unchanged. An entirely blank response creates no row.
6. Direct changes to stock quantities or dates refresh the dashboard. A header
   edit triggers a manifest sync, restoring generated headers for manifest items.

Do not add or delete stock columns manually: use the manifest. Unrecognised
headers are rejected instead of being adopted as items. No Sheets protections
are installed. Inactive or removed items retain their stock columns and labels;
they disappear from the form, not from historical stock data. Re-enabling an ID
reuses its retained column. Physical deletion of columns is not a supported
method of retiring items or a recoverable undo operation.

Google Forms still assigns its own question IDs. Those IDs are internal
references; inventory identity comes exclusively from ID Barang.

Local mocks cannot verify deployed triggers or real Sheets/Form permissions.
After deployment, test a rename, deactivation, form submission with a zero and
a blank answer, and a direct stock correction using disposable data.
