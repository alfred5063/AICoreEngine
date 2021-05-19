from FlaskRESTAPI import db

tasks = db.Table(
  'tasks',
  db.metadata,
  autoload=True,
  autoload_with=db.engine
)
