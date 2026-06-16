from fastapi import APIRouter, Depends

from app.deps import DbSession, StoreUser, require_role
from app.enums import Role
from app.schemas.salespeople import SalespersonRead
from app.services import salespeople as svc

router = APIRouter(prefix="/salespeople", tags=["salespeople"])

_BARISTA_PLUS = require_role(Role.OWNER, Role.MANAGER, Role.BARISTA, Role.BAKER)


@router.get(
    "",
    response_model=list[SalespersonRead],
    operation_id="salespeople_list",
    dependencies=[Depends(_BARISTA_PLUS)],
)
async def list_salespeople(user: StoreUser, db: DbSession) -> list[SalespersonRead]:
    return await svc.list_salespeople(db=db, store_id=user.store_id)
