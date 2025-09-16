# models.py
from extensions import db
from datetime import datetime

# ... (test_case_tags 和 Tag 模型的定義不變) ...
test_case_tags = db.Table('test_case_tags',
    db.Column('test_case_id', db.Integer, db.ForeignKey('test_case.id'), primary_key=True),
    db.Column('tag_id', db.Integer, db.ForeignKey('tag.id'), primary_key=True)
)

class Tag(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), unique=True, nullable=False)

    def __repr__(self):
        return f'<Tag {self.name}>'


class TestCase(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    product_type = db.Column(db.String(50), nullable=False, default='未分類產品')
    category = db.Column(db.String(100), nullable=False)
    main_category = db.Column(db.String(50), nullable=True)
    sub_category = db.Column(db.String(50), nullable=True)
    case_id = db.Column(db.String(50), unique=True, nullable=False)
    test_item = db.Column(db.String(200), nullable=False)
    test_purpose = db.Column(db.Text, nullable=True)
    preconditions = db.Column(db.Text, nullable=True)
    test_steps = db.Column(db.Text, nullable=True)
    expected_result = db.Column(db.Text, nullable=True)
    actual_result = db.Column(db.Text, nullable=True)
    status = db.Column(db.String(20), nullable=False, default='未執行')
    notes = db.Column(db.Text, nullable=True)
    reference = db.Column(db.String(200), nullable=True)

    tags = db.relationship('Tag', secondary=test_case_tags, lazy='subquery',
                           backref=db.backref('test_cases', lazy=True))

    # ★★★ 核心修正點 1: 新增與 Attachment 的關聯 ★★★
    attachments = db.relationship('Attachment', backref='test_case', lazy=True, cascade="all, delete-orphan")


# ★★★ 核心修正點 2: 新增 Attachment 模型 ★★★
class Attachment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(255), nullable=False)
    filepath = db.Column(db.String(255), nullable=False) # 相對於 ATTACHMENT_FOLDER 的路徑
    uploaded_on = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    test_case_id = db.Column(db.Integer, db.ForeignKey('test_case.id'), nullable=False)

    def __repr__(self):
        return f'<Attachment {self.filename}>'