from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.errors import NotFound
from app.models.sales import Salesperson
from app.schemas.salespeople import SalespersonRead


async def list_salespeople(db: AsyncSession, *, store_id: str) -> list[SalespersonRead]:
    rows = (
        (
            await db.execute(
                select(Salesperson)
                .where(Salesperson.store_id == store_id, Salesperson.is_active.is_(True))
                .order_by(Salesperson.name)
            )
        )
        .scalars()
        .all()
    )
    return [SalespersonRead.model_validate(r) for r in rows]


async def get_salesperson(db: AsyncSession, *, store_id: str, sales_id: str) -> Salesperson:
    sp = (
        (
            await db.execute(
                select(Salesperson).where(
                    Salesperson.id == sales_id,
                    Salesperson.store_id == store_id,
                    Salesperson.is_active.is_(True),
                )
            )
        )
        .scalars()
        .first()
    )
    if not sp:
        raise NotFound("Salesperson not found")
    return sp
