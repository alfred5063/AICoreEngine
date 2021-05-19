from FlaskRESTAPI import db

jobs = db.Table(
  'jobs',
  db.metadata,
  autoload=True,
  autoload_with=db.engine
)
