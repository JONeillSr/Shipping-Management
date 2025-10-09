Subject: {{email_subject}}

Hello {{auction_info.logistics_contact.email}},

Please provide a freight quote for the following:

**Auction / Pickup**
- Auction: {{auction_info.auction_name}}
- Pickup window: {{auction_info.pickup_datetime}}
- Pickup address:
  {{auction_info.pickup_address}}

**Delivery**
- Address: {{delivery_address}}
- Delivery window: {{auction_info.delivery_datetime}}
- Notice: {{auction_info.delivery_notice}}

**On-site / Special Notes**
{{auction_info.special_notes}}

**Shipping Requirements**
- Labor: {{shipping_requirements.labor_needed}}
- Trucks: {{shipping_requirements.truck_types}}
- Total pallets: {{shipping_requirements.total_pallets}}
- Weight notes: {{shipping_requirements.weight_notes}}

Please include:
- All-in cost (pickup → delivery)
- Accessorials (liftgate, tarp, detention, etc.)
- Lead time and transit time
- Insurance details

Thanks,
{{_meta.requester_name}}  
{{_meta.requester_phone}}
