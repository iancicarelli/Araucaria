class Idlocals:
    def __init__(self, id, name):
        self.id = id
        self.name = name

    def __str__(self):
        return f"Locals(code={self.id}, name={self.name})"