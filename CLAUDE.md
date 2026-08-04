# Finocam · Dashboard Amazon Vendor

Eres mi asistente de trabajo en Roicos. Gestiono reporting para clientes de Amazon.

## Proyecto activo: Finocam · Cabero Group (Amazon Vendor)
Fichero principal: C:\Users\Daniela\Downloads\FINOCA_2_catTable.HTM
Drive file ID: 1LJArMKZZYrNGcmTozWUN8MF2MQ3pa8xU
Repo GitHub: https://github.com/dparodi-roicos/Finocam-Amazon
URL pública base: https://dparodi-roicos.github.io/Finocam-Amazon/
Ruta local repo Git: C:\Users\Daniela\Desktop\Git Finocam\

## Fuentes de datos
- Ventas Vendor (sell-out/sell-in) → MCP MerchantSpring (marketplace: amazon_vendor)
  - Sell-out = orderedRevenue
  - Sell-in  = shippedRevenueManufacturing (vista manufacturing siempre, todos los mercados)
- Publicidad (SP/SB/SD) → MCP Pacvue (retailer: amazon-ads)
- Meta Ads → MCP Metricool (brandId 2216922)
- Regla de integridad: NO inventar nada — si no hay dato se pone '-'

## IDs MerchantSpring Finocam (Amazon Vendor · 9 mercados)
- ES: channelId 21907968  · merchantId "amzn1.vg.3250111 @ A1RKKUPIHCS9HS"
- FR: channelId 22527353  · merchantId "amzn1.vg.6776082 @ A13V1IB3VIYZZH"
- IT: channelId 22527944  · merchantId "amzn1.vg.6123122 @ APJ6JRA9NG5V4"
- DE: channelId 23167335  · merchantId "amzn1.vg.6776022 @ A1PA6795UKMFR9"
- NL: channelId 23167955  · merchantId "amzn1.vg.6968302 @ A1805IZSGTT6HS"
- PL: channelId 30609951  · merchantId "amzn1.vg.6968312 @ A1C3SOZRARQ6R3"
- UK: channelId 78652228  · merchantId "amzn1.vg.6776072 @ A1F83G8C2ARO7P"
- BE: channelId 80601759  · merchantId "amzn1.vg.8494352 @ AMEN7PMS3EDWL"
- SE: channelId 108423884 · merchantId "amzn1.vg.6968292 @ A2NODRKZP88ZB9" (sin ads)

## Perfiles Pacvue Finocam (amazon-ads · 8 mercados, SE sin ads)
- ES: id 1282307461612898 · "Finocam [ES][ES]"
- FR: id 4335491460860834 · "Finocam[FR]"
- IT: id 3083150633499829 · "Finocam[IT]"
- DE: id 1911014084655388 · "Finocam[DE]"
- NL: id 2830976541696187 · "Finocam[NL]"
- UK: id 2656163307834937 · "Finocam[UK]"
- BE: id 2269833629686159 · "Finocam[BE]"
- PL: id 34385048011383   · "Finocam[PL]"

## Publicar en GitHub (siempre estos 4 pasos en PowerShell)
1. Copy-Item "C:\Users\Daniela\Downloads\FINOCA_2_catTable.HTM" "C:\Users\Daniela\Desktop\Git Finocam\FINOCA_2_catTable.HTM" -Force
2. cd "C:\Users\Daniela\Desktop\Git Finocam"
3. git add FINOCA_2_catTable.HTM ; git commit -m "[descripción]"
4. git push

## Guardar en Drive
Copy-Item "C:\Users\Daniela\Downloads\FINOCA_2_catTable.HTM" "G:\.shortcut-targets-by-id\1qUjcAS06hhm_53qObrDfg5Yw3Vr2xQ_3\FINOCA_2_catTable.HTM" -Force
