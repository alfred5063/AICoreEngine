import datetime
from FlaskRESTAPI import db

process_output = db.Table(
  'process_outputs',
  db.Column('id', db.Integer, primary_key=True),
  db.Column('task_id', db.Integer, db.ForeignKey('tasks.id')),
  db.Column('guid', db.String),
  db.Column('task_info', db.String),
  db.Column('status', db.String),
  db.Column('filename', db.String),
  db.Column('timestamp', db.DateTime, default=datetime.datetime.utcnow),
)

