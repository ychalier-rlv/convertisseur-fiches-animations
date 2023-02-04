class AnimationMetadata:

	KEYS = {
		"thématiques": "topics",
		"thématique": "topics",
		"participants": "participants",
		"durée": "duration",
		"public": "audience",
		"prérequis": "prerequisites",
		"matériel": "material",
	}

	def __init__(self):
		self.participants = None
		self.duration = None
		self.audience = None
		self.prerequisites = None
		self.material = []
		self.topics = []
	
	@classmethod
	def from_dict(cls, d):
		o = cls()
		o.participants = d.get("participants")
		o.duration = d.get("duration")
		o.audience = d.get("audience")
		o.prerequisites = d.get("prerequisites")
		o.material = d.get("material", [])
		o.topics = d.get("topics", [])
		return o
	
	def to_dict(self):
		return {
			"participants": self.participants,
			"duration": self.duration,
			"audience": self.audience,
			"prerequisites": self.prerequisites,
			"material": self.material,
			"topics": self.topics,
		}


class AnimationStep:

	def __init__(self):
		self.title = None
		self.duration = None
		self.content = None
	
	@classmethod
	def from_dict(cls, d):
		o = cls()
		o.title = d.get("title")
		o.duration = d.get("duration")
		o.content = d.get("content")
		return o
	
	def to_dict(self):
		return {
			"title": self.title,
			"duration": self.duration,
			"content": self.content,
		}


class Animation:

	def __init__(self):
		self.title = None
		self.subtitle = None
		self.steps = []
		self.others = {}
		self.metadata = AnimationMetadata()
		self.resources = []

	@classmethod
	def from_dict(cls, d):
		o = cls()
		o.title = d.get("title")
		o.subtitle = d.get("subtitle")
		o.others = d.get("others", {})
		o.steps = [AnimationStep.from_dict(dd) for dd in d.get("steps", [])]
		o.metadata = AnimationMetadata.from_dict(d.get("metadata"))
		o.resources = d.get("resources", [])
		return o
	
	def to_dict(self):
		return {
			"title": self.title,
			"subtitle": self.subtitle,
			"steps": [step.to_dict() for step in self.steps],
			"metadata": self.metadata.to_dict(),
			"others": self.others,
			"resources": self.resources
		}