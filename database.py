from sqlalchemy.orm import *
from sqlalchemy import *

ENGINE = create_engine(r'sqlite:///C:\Users\ubuntu\Documents\Отдел корпоративной защиты\personal.db', echo=True)

Base = declarative_base()


class Personal(Base):
    __abstract__ = True


class Candidate(Personal):
    """ Create model for candidates dates"""

    __tablename__ = 'candidates'

    id = Column(Integer, nullable=False, unique=True, primary_key=True, autoincrement=True)
    staff = Column(Text)
    department = Column(Text)
    full_name = Column(Text, index=True)
    last_name = Column(Text)
    birthday = Column(Text)
    birth_place = Column(Text)
    country = Column(Text)
    series_passport = Column(Text)
    number_passport = Column(Text)
    date_given = Column(Text)
    snils = Column(Text)
    inn = Column(Text)
    reg_address = Column(Text)
    live_address = Column(Text)
    phone = Column(Text)
    email = Column(Text)
    education = Column(Text)
    check = relationship('Check', back_populates='candidate')
    inquery = relationship('Inquery', back_populates='candidate')
    registr = relationship('Registr', back_populates='candidate')


class Check(Personal):
    """ Create model for candidates checks"""

    __tablename__ = 'checks'

    id = Column(Integer, nullable=False, unique=True, primary_key=True, autoincrement=True)
    check_work_place = Column(Text)
    check_passport = Column(Text)
    check_debt = Column(Text)
    check_bankruptcy = Column(Text)
    check_bki = Column(Text)
    check_affiliation = Column(Text)
    check_internet = Column(Text)
    check_cronos = Column(Text)
    check_cross = Column(Text)
    resume = Column(Text)
    date_check = Column(Text)
    officer = Column(Text)
    check_id = Column(Integer, ForeignKey('candidates.id'))
    candidate = relationship('Candidate', back_populates='check')


class Inquery(Personal):
    """ Create model for candidates iqueries"""

    __tablename__ = 'iqueries'

    id = Column(Integer, nullable=False, unique=True, primary_key=True, autoincrement=True)
    staff = Column(Text)
    period = Column(Text)
    info = Column(Text)
    firm = Column(Text)
    date_inq = Column(Text)
    iquery_id = Column(Integer, ForeignKey('candidates.id'))
    candidate = relationship('Candidate', back_populates='inquery')


class Registr(Personal):
    """ Create model for candidates iqueries"""

    __tablename__ = 'registries'

    id = Column(Integer, nullable=False, unique=True, primary_key=True, autoincrement=True)
    checks = Column(Text)
    recruiter = Column(Text)
    final_date = Column(Text)
    url = Column(Text)
    registry_id = Column(Integer, ForeignKey('candidates.id'))
    candidate = relationship('Candidate', back_populates='registr')


# Personal.metadata.create_all(bind=ENGINE)
