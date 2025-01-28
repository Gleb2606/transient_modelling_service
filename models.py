from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()

class Model(db.Model):
    __tablename__ = 'models'
    id = db.Column(db.Integer, primary_key=True)
    data_in_frame = db.Column(db.Integer, nullable=False)
    input_width = db.Column(db.Integer, nullable=False)
