class Format:
    def __init__(self, code, amount,name_id):
        self.code = code
        self.amount = amount
        self.name_id = name_id

    def __str__(self):
        return f"Format(code={self.code}, amount={self.amount},id={self.name_id})"
