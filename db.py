from sqlalchemy import Column, Integer, JSON, String, Float, Boolean, ForeignKey, create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import relationship, sessionmaker

MAX_LESSONS = 10

Base = declarative_base()


class User(Base):
    __tablename__ = 'user'

    id = Column(Integer, primary_key=True)
    telegram_id = Column(Integer, unique=True)
    subjects = relationship('Subject', cascade='all, delete-orphan')
    schedule = Column(JSON)
    settings = relationship('Settings', uselist=False, back_populates='user')
    homework = relationship('Homework', cascade='all, delete-orphan')

    def __repr__(self):
        return 'User with telegram ID {}'.format(self.telegram_id)


class Subject(Base):
    __tablename__ = 'subject'

    id = Column(Integer, primary_key=True)
    name = Column(String, nullable=False)
    teacher = Column(String)
    room = Column(String)

    user_id = Column(Integer, ForeignKey('user.id'))

    def __repr__(self):
        return 'Subject "{}"'.format(self.name)


class Settings(Base):
    __tablename__ = 'settings'

    id = Column(Integer, primary_key=True)
    language = Column(String, nullable=False)
    timezone = Column(Float, nullable=False)

    user_id = Column(Integer, ForeignKey('user.id'))
    user = relationship('User', back_populates='settings')

    def __repr__(self):
        return 'Language: "{}", timezone: {}'.format(self.language, self.timezone)


class Homework(Base):
    __tablename__ = 'homework'

    id = Column(Integer, primary_key=True)
    date = Column(Integer, nullable=False)
    subject = Column(Integer, nullable=False)
    for_lesson = Column(Boolean, nullable=False)
    description = Column(String, nullable=False)

    user_id = Column(Integer, ForeignKey('user.id'))

    def __repr__(self):
        return 'Homework "{}"'.format(self.description)


engine = create_engine('sqlite:///files/db.sqlite3', echo=True)
Base.metadata.create_all(engine)

Session = sessionmaker(bind=engine)


def get_user(user_id):
    session = Session()
    q = session.query(User).filter_by(telegram_id=user_id)
    if not session.query(q.exists()).scalar():
        settings = Settings(language='en', timezone=3)
        user = User(telegram_id=user_id, settings=settings)
        session.add(user)
        session.commit()
    else:
        user = q.first()

    return session, user
