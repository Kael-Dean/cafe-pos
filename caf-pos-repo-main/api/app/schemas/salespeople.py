from pydantic import BaseModel


class SalespersonRead(BaseModel):
    model_config = {"from_attributes": True}

    id: str
    name: str
    is_active: bool


class AssignSalesRequest(BaseModel):
    sales_id: str | None = None
