class Order:
    def __init__(self, retailer_id, date_required=None, install_date=None, first_poll_date=None,
                 service_activated=None, message=None, body=None):
        self.retailer_id = retailer_id  # Attribute for Retailer ID
        self.date_required = date_required  # Optional attribute 'O'
        self.install_date = install_date  # Optional attribute 'P'
        self.first_poll_date = first_poll_date  # Optional attribute 'Q'
        self.service_activated = service_activated  # Optional attribute 'R'
        self.message = message
        self.body = body
