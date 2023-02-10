import argparse
import glob
import json
import os
import re
import shutil
import unicodedata
import urllib.parse
import xml.etree.ElementTree

import docx
import tqdm


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

	def to_markdown(self):
		md = ""
		metadata_keys = {
			"topics": "Thématiques",
			"participants": "Participants",
			"duration": "Durée",
			"audience": "Public",
			"prerequisites": "Prérequis",
			"material": "Matériel"
		}
		for metadata_key, label in metadata_keys.items():
			value = getattr(self, metadata_key)
			if value is None:
				continue
			md += f"**{ label }**\n"
			if isinstance(value, list):
				for value_item in value:
					md += f"- { value_item }\n"
				md += "\n"
			else:
				md += f"{ value }\n\n"
		return md


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
	
	def to_markdown(self):
		md = ""
		if self.title is not None:
			md += f"### { self.title } ({ self.duration } min)\n\n"
		md += f"{ self.content }\n\n"
		return md


class Animation:

	def __init__(self):
		self.title = None
		self.subtitle = None
		self.steps = []
		self.others = {}
		self.metadata = AnimationMetadata()
		self.resources = []
		self.online_resources = []

	@classmethod
	def from_dict(cls, d):
		o = cls()
		o.title = d.get("title")
		o.subtitle = d.get("subtitle")
		o.others = d.get("others", {})
		o.steps = [AnimationStep.from_dict(dd) for dd in d.get("steps", [])]
		o.metadata = AnimationMetadata.from_dict(d.get("metadata"))
		o.resources = d.get("resources", [])
		o.online_resources = d.get("online_resources", [])
		return o
	
	def to_dict(self):
		return {
			"title": self.title,
			"subtitle": self.subtitle,
			"steps": [step.to_dict() for step in self.steps],
			"metadata": self.metadata.to_dict(),
			"others": self.others,
			"resources": self.resources,
			"online_resources": self.online_resources,
		}
	
	def to_markdown(self):
		md = ""
		md += f"# { self.title }\n\n"
		if self.subtitle is not None:
			md += f"{ self.subtitle }\n\n"
		md += self.metadata.to_markdown()
		for key, content in self.others.items():
			md += f"## { key }\n\n{ content }\n\n"
		md += "## Déroulé\n\n"
		for step in self.steps:
			md += step.to_markdown()
		if self.resources:
			md += "## Ressources\n\n"
			for res in self.resources:
				md += f"- [{ res['name'] } ({ res['ext'][1:].upper() }, { sizeof_fmt(res['size']) })]({ urllib.parse.quote(res['slug'] + res['ext']) })\n"
			md += "\n"
		if self.online_resources:
			md += "## Ressources en ligne\n\n"
			for res in self.online_resources:
				md += f"- [{ res['name'] }]({ res['url'] })\n"
			md += "\n"
		return md


def sizeof_fmt(num, suffix="o"):
	for unit in ["", "K", "M", "G", "T", "P", "E", "Z"]:
		if abs(num) < 1024.0:
			if num == int(num):
				return f"{num} {unit}{suffix}"
			if abs(num) >= 100:
				return f"{num:3.0f} {unit}{suffix}"
			return f"{num:2.1f} {unit}{suffix}"
		num /= 1024.0
	return f"{num:.1f}Yi{suffix}"


def docx_convert_paragraph_text_to_markdown(paragraph):
	text = ""
	root = xml.etree.ElementTree.fromstring(paragraph._element.xml)
	NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
	NSR = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
	for child in root:
		if child.tag == NS + "r":
			child_is_bold = False
			child_is_italic = False
			child_text = ""
			for sub_child in child:
				if sub_child.tag == NS + "t":
					child_text += sub_child.text
				elif sub_child.tag == NS + "rPr":
					for sub_sub_child in sub_child:
						if sub_sub_child.tag == NS + "b":
							child_is_bold = True
						elif sub_sub_child.tag == NS + "i":
							child_is_italic = True
			if child_is_bold and child_is_italic:
				text += f"**_{ child_text.strip() }_** "
			elif child_is_bold:
				text += f"__{ child_text.strip() }__ "
			elif child_is_italic:
				text += f"_{ child_text.strip() }_ "
			else:
				text += child_text.strip() + " "
		elif child.tag == NS + "hyperlink":
			link_text = child.find(NS + "r").find(NS + "t").text
			link_url = paragraph.part.rels[child.attrib[NSR + "id"]].target_ref
			text += f"[{ link_text }]({ link_url }) "
	text = re.sub(r"([\(\[]) ", r"\1", re.sub(r" ([,\.\)\]])", r"\1", text))
	return text


def docx_convert_to_markdown(*paragraphs):
	md = ""
	start_of_list = True
	for paragraph in paragraphs:
		if paragraph.style.name == "List Paragraph":
			if start_of_list:
				md += "\n"
			md += "\n- " + docx_convert_paragraph_text_to_markdown(paragraph)
			start_of_list = False
		else:
			md += "\n\n" + docx_convert_paragraph_text_to_markdown(paragraph)
			start_of_list = True
	return md.strip()


def get_directory_size(path="."):
    total_size = 0
    for dirpath, dirnames, filenames in os.walk(path):
        for f in filenames:
            fp = os.path.join(dirpath, f)
            if not os.path.islink(fp):
                total_size += os.path.getsize(fp)
    return total_size


class DocumentParser:

	def __init__(self, path):
		self.path = path
		self.document = docx.Document(path)
		self.animation = Animation()
		self.section = None
	
	def parse_metadata_section(self):
		previous_key = None 
		values_count = 0
		for paragraph in self.section:
			if paragraph.text.lower() in AnimationMetadata.KEYS:
				previous_key = AnimationMetadata.KEYS[paragraph.text.lower()]
				values_count = 0
			elif previous_key is not None:
				values_count += 1
				if values_count == 1:
					setattr(self.animation.metadata, previous_key, paragraph.text)
				elif values_count == 2:
					setattr(self.animation.metadata, previous_key, [getattr(self.animation.metadata, previous_key), paragraph.text])
				else:
					setattr(self.animation.metadata, previous_key, getattr(self.animation.metadata, previous_key) + [paragraph.text])		

	def parse_section(self):
		if self.section[0].text == "Déroulé":
			self.animation.steps = []
			i, j = 1, 1
			while i < len(self.section) - 1:
				while i < len(self.section) - 1 and self.section[i + 1].style.name != "Heading 2":
					i += 1
				step = AnimationStep()
				if self.section[j].style.name == "Heading 2":
					match = re.search("^(.+) (?:\((\d+) min(?:utes?)?\))?$", self.section[j].text.strip())
					if match is None:
						continue
					step.title = match.group(1)
					step.duration = match.group(2)
					step.content = docx_convert_to_markdown(*self.section[j + 1:i + 1])
				else:
					step.content = docx_convert_to_markdown(*self.section[j:i + 1])
				self.animation.steps.append(step)
				j = i + 1
				i += 1
		else:
			self.animation.others[self.section[0].text] = docx_convert_to_markdown(*self.section[1:])	
	
	def add_resources_file(self, path):
		split = os.path.splitext(os.path.basename(path))
		self.animation.resources.append({
			"name": split[0],
			"type": "file",
			"ext": split[1],
			"size": os.path.getsize(path),
			"path": os.path.realpath(path),
			"slug": slugify(split[0]),
		})
	
	def add_resources_folder(self, path):
		self.animation.resources.append({
			"name": os.path.basename(path),
			"type": "folder",
			"ext": ".zip",
			"size": get_directory_size(path),
			"path": os.path.realpath(path),
			"slug": slugify(os.path.basename(path)),
		})

	def add_resources_url(self, path):
		url = None
		with open(path, "r", encoding="utf8") as file:
			for line in file.read().split("\n"):
				if line[:4] == "URL=":
					url = line[4:].strip()
		if url is None:
			print("Could not get URL from", os.path.realpath(path))
			return	
		self.animation.online_resources.append({
			"name": os.path.splitext(os.path.basename(path))[0],
			"url": url,
		})

	def parse_resources(self):
		for path in glob.glob(os.path.join(os.path.dirname(self.path), "*")):
			if path == self.path:
				continue
			if os.path.splitext(path)[1] == ".lnk":
				continue
			elif os.path.splitext(path)[1] == ".url":
				self.add_resources_url(path)
			elif os.path.isfile(path):
				self.add_resources_file(path)
			elif os.path.isdir(path):
				self.add_resources_folder(path)

	def parse(self):
		first_section = True
		self.section = []
		for paragraph in self.document.paragraphs:
			if paragraph.style.name == "Title":
				self.animation.title = paragraph.text
			elif paragraph.style.name == "Subtitle":
				self.animation.subtitle = paragraph.text
			elif paragraph.style.name == "Heading 1":
				if first_section:
					first_section = False
					self.parse_metadata_section()
				else:
					self.parse_section()
				self.section = [paragraph]
			else:
				self.section.append(paragraph)
		self.parse_section()
		self.parse_resources()
		return self.animation


def windows_safe_filename(string):
	"""Convert a string to a valid Windows file name.
	"""
	return re.sub(" +", " ", re.sub(r"/<>:\"\\\|\?\*", "", string))


def strip_accents(s):
   return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')


def slugify(string):
	return re.sub(r'[\W_]+', '-', strip_accents(re.sub("’", "", string.lower())))


def generate_animation_output(animation, output_folder, copy_resources=True):
	os.makedirs(output_folder, exist_ok=True)
	animation_folder = os.path.join(output_folder, windows_safe_filename(slugify(animation.title)))
	if os.path.isdir(animation_folder):
		shutil.rmtree(animation_folder)
	os.makedirs(animation_folder, exist_ok=False)
	with open(os.path.join(animation_folder, "default.fr.md"), "w", encoding="utf8") as file:
		file.write(f"---\ntitle: \"{ animation.title }\"\n---\n\n")
		file.write(animation.to_markdown())
	if not copy_resources:
		return
	for res in animation.resources:
		if res["type"] == "file":
			shutil.copy(res["path"], os.path.join(animation_folder, res["slug"] + res["ext"]))
		elif res["type"] == "folder":
			shutil.make_archive(os.path.join(animation_folder, res["slug"]), "zip", res["path"])


def find_animation_paths(top):
	if os.path.isfile(top):
		return [top]
	if top.endswith("\""):
		top = top[:-1]
	paths = []
	for root, dirs, files in os.walk(top, topdown=True):
		if "ATELIERS IDÉES" in root:
			continue
		for f in files:
			if f == "Fiche animation.docx":
				paths.append(os.path.join(root, f))
	return paths


def main():
	parser = argparse.ArgumentParser()
	parser.add_argument("input_path", type=str, help="path to a DOCX file or a folder")
	parser.add_argument("-o", "--output_path", type=str, help="path to the output folder", default="animations")
	parser.add_argument("-n", "--no-copy", action="store_true", help="do not copy resources")
	args = parser.parse_args()		
	db_path = os.path.join(args.output_path, "index.json")
	db = {}
	if os.path.isfile(db_path):
		with open(db_path, "r", encoding="utf8") as file:
			db = json.load(file)
	for animation_path in tqdm.tqdm(find_animation_paths(args.input_path), unit="file"):
		animation = DocumentParser(animation_path).parse()
		db[animation_path] = animation.to_dict()
		generate_animation_output(animation, args.output_path, copy_resources=not args.no_copy)
	with open(db_path, "w", encoding="utf8") as file:
		json.dump(db, file, indent=4, default=str)


if __name__ == "__main__":
	main()
