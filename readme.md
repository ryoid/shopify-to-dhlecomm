Export orders from Shopify as CSV

Specific orders:
Shopify -> orders
- Tick orders to export
- Top right "Export"
- Tick "Selected: n orders"
- Export
- Move file in this folder (recommended)

Ensure you have "DHLeC.xlsx" (Template), found in Sample Template.

Required fields
- account (Pick-up Account Number)
- serviceCode (Shipping Service Code)
- description (Shipment Description)
- unitWeight (Used to multiply by quantity for weight)

Caveats
Shopify order export does not have certain info, like Weight and HS Code so these values are supplied manually.

Example
./shopify-to-dhlecomm convert --account 123456 --salesChannel Website --serviceCode PLT --originCountry SG --unitWeight 20 --description "Mechanical Keyboard Components"  --incoterm DDP "your order exports.csv" "dhl_template.xlsx"

./shopify-to-dhlecomm convert --account 123456 --salesChannel Website --serviceCode PLT --originCountry SG --unitWeight 20 --description "Mechanical Keyboard Components"  --incoterm DDP ".\orders_export.csv" DHLeC.xlsx

Exports to "output-[time].xlsx"