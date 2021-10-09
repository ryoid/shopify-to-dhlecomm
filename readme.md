# Shopify-to-DHLecomm

Export orders from Shopify as CSV and convert them to DHL eCommerce bulk upload format.

## Usage
### Exporting orders
Specific orders:
Shopify -> orders
- Tick orders to export
- Top right "Export"
- Tick "Selected: n orders"
- Export
- Move file in this folder (recommended)

### Run

#### Required fields
- account (Pick-up Account Number)
- serviceCode (Shipping Service Code)
- description (Shipment Description)
- unitWeight (Used to multiply by quantity for weight)

#### Caveats
Shopify order export does not have certain info, like Weight and HS Code so these values are supplied manually.

#### Example
```sh
./shopify-to-dhlecomm convert --account 123456 --salesChannel Website --serviceCode PLT --originCountry SG --unitWeight 20 --description "Clothes"  --incoterm DDP "your order exports.csv" ".\Sample Template\DHLeC AP Portal upload file_1.0_en.xlsx"
```

```sh
./shopify-to-dhlecomm convert --account 123456 --salesChannel Website --serviceCode PLT --originCountry SG --unitWeight 20 --description "Clothes"  --incoterm DDP ".\orders_export.csv" ".\Sample Template\DHLeC AP Portal upload file_1.0_en.xlsx"
```

Exports to "output-[time].xlsx"