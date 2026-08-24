"""认证路由：注册 / 登录 / 当前用户。"""
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session

from backend.auth import hash_password, verify_password, create_access_token, get_current_user
from backend.database import get_db
from backend.models import User, Tenant
from backend.schemas import RegisterIn, LoginIn, TokenOut, UserOut

router = APIRouter(prefix="/api/auth", tags=["auth"])


def _user_to_out(u: User, tenant_name: str) -> UserOut:
    return UserOut(
        id=u.id, email=u.email, name=u.name,
        tenant_id=u.tenant_id, tenant_name=tenant_name,
        is_active=u.is_active, created_at=u.created_at,
    )


@router.post("/register", response_model=TokenOut)
def register(payload: RegisterIn, db: Session = Depends(get_db)):
    if db.query(User).filter(User.email == payload.email).first():
        raise HTTPException(status_code=400, detail="该邮箱已注册")

    tenant_name = payload.tenant_name or payload.email.split("@")[-1]
    tenant = Tenant(name=tenant_name)
    db.add(tenant)
    db.flush()

    user = User(
        tenant_id=tenant.id,
        email=payload.email,
        name=payload.name,
        hashed_password=hash_password(payload.password),
    )
    db.add(user)
    db.commit()
    db.refresh(user)

    token = create_access_token(user.id, user.tenant_id)
    return TokenOut(access_token=token, user=_user_to_out(user, tenant.name))


@router.post("/login", response_model=TokenOut)
def login(payload: LoginIn, db: Session = Depends(get_db)):
    user = db.query(User).filter(User.email == payload.email).first()
    if not user or not verify_password(payload.password, user.hashed_password):
        raise HTTPException(status_code=401, detail="邮箱或密码错误")
    if not user.is_active:
        raise HTTPException(status_code=403, detail="账号已禁用")

    tenant = db.query(Tenant).filter(Tenant.id == user.tenant_id).first()
    token = create_access_token(user.id, user.tenant_id)
    return TokenOut(access_token=token, user=_user_to_out(user, tenant.name if tenant else ""))


@router.get("/me", response_model=UserOut)
def me(current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    tenant = db.query(Tenant).filter(Tenant.id == current_user.tenant_id).first()
    return _user_to_out(current_user, tenant.name if tenant else "")
