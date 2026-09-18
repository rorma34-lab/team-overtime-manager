from pydantic import BaseModel, Field
from typing import Optional, List

class UserLoginRequest(BaseModel):
    emp_id: str = Field(..., description="사원번호")

class UserRegisterRequest(BaseModel):
    emp_id: str = Field(..., description="사원번호")
    name: str = Field(..., description="이름")
    team: str = Field(..., description="소속팀")
    position: Optional[str] = Field("팀원", description="직급")
    is_admin: Optional[int] = Field(0, description="관리자 여부(0 또는 1)")

class UserUpdateRequest(BaseModel):
    name: Optional[str] = None
    team: Optional[str] = None
    position: Optional[str] = None
    is_admin: Optional[int] = None
    is_super: Optional[int] = None
    admin_emp_id: Optional[str] = None

class OvertimeCreateRequest(BaseModel):
    emp_id: str = Field(..., description="신청자 사원번호")
    category: str = Field(..., description="특근 분류: 대체근무 / 일반휴일 / 법정휴일 / 대체휴무")
    start_date: str = Field(..., description="시작일 (YYYY-MM-DD)")
    end_date: str = Field(..., description="종료일 (YYYY-MM-DD)")
    project_no: Optional[str] = Field("", description="프로젝트 번호")
    location: Optional[str] = Field("", description="장소")
    reason: Optional[str] = Field("", description="특근 사유")
    sub_holiday_used: Optional[float] = Field(0.0, description="대체휴무 사용일수")
    sub_holiday_date: Optional[str] = Field("", description="대체휴무 사용일자 (YYYY-MM-DD)")
    bonus_granted: Optional[int] = Field(0, description="보너스 부여 여부 (1: 부여, 0: 미부여)")
    is_pre_deduct: Optional[int] = Field(0, description="사전차감 여부 (1: 사전차감, 0: 일반)")
    trip_start_date: Optional[str] = Field("", description="대체휴무 출장기간 시작일 (YYYY-MM-DD)")
    trip_end_date: Optional[str] = Field("", description="대체휴무 출장기간 종료일 (YYYY-MM-DD)")

class OvertimeUpdateRequest(BaseModel):
    changed_by: str = Field(..., description="수정 작업을 수행하는 사원번호")
    category: str = Field(..., description="특근 분류")
    start_date: str = Field(..., description="시작일 (YYYY-MM-DD)")
    end_date: str = Field(..., description="종료일 (YYYY-MM-DD)")
    project_no: Optional[str] = ""
    location: Optional[str] = ""
    reason: Optional[str] = ""
    sub_holiday_used: Optional[float] = 0.0
    sub_holiday_date: Optional[str] = ""
    bonus_granted: Optional[int] = None
    is_pre_deduct: Optional[int] = None
    trip_start_date: Optional[str] = ""
    trip_end_date: Optional[str] = ""

class OvertimeDeleteRequest(BaseModel):
    changed_by: str = Field(..., description="삭제 작업을 수행하는 사원번호")

class OvertimeConfirmRequest(BaseModel):
    admin_emp_id: str = Field(..., description="확인 처리하는 관리자 사번")
    is_confirmed: int = Field(..., description="1: 확인/승인, 0: 취소/대기")

class OvertimePreDeductRequest(BaseModel):
    admin_emp_id: str = Field(..., description="사전차감 토글하는 사원 또는 관리자 사번")
    is_pre_deduct: int = Field(..., description="1: 사전차감 ON, 0: 사전차감 OFF")

class ExportRequest(BaseModel):
    ids: Optional[List[int]] = None
    start_date: Optional[str] = None
    end_date: Optional[str] = None
    team: Optional[str] = None
    category: Optional[str] = None
    is_confirmed: Optional[int] = None
    search: Optional[str] = None
    admin_emp_id: Optional[str] = None

class TeamCreateRequest(BaseModel):
    name: str = Field(..., description="소속팀 명칭")
    admin_emp_id: str = Field(..., description="슈퍼관리자 사원번호")

class BackupSaveRequest(BaseModel):
    name: Optional[str] = Field(None, description="백업 명칭")
    label: Optional[str] = Field(None, description="백업 설명 라벨")
    description: Optional[str] = Field("", description="백업 상세 설명")
    admin_emp_id: Optional[str] = Field(None, description="관리자 사원번호")

class BackupLoadRequest(BaseModel):
    filename: str = Field(..., description="복원할 백업 파일명")
    admin_emp_id: Optional[str] = Field(None, description="슈퍼관리자 사원번호")


