import argparse
import glob
import json
import os
import re
import shutil
import xml.etree.ElementTree as ET

import docx
import tqdm
from animation import Animation, AnimationMetadata, AnimationStep


def convert_paragraph_text_to_markdown(paragraph):
	text = ""
	root = ET.fromstring(paragraph._element.xml)
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
	return text
	print("=" * 80)
	# for link in paragraph._element.xpath(".//w:hyperlink"):
	# 	inner_run = link.xpath("w:r", namespaces=link.nsmap)[0]
	# 	rId = link.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
	# 	print(inner_run.text, rId, paragraph.part.rels[rId]._target)
	for prun in paragraph.runs:
		# print(prun.text)
		# for relId, rel in prun.part.rels.items():
		# 	if rel.reltype == docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK:
		# 		print(relId)
		# 		print(rel.target_ref)
		# 		print(dir(rel))
		print("-" * 10)
		print(prun._element.xml)
		# for link in prun._element.xpath("w:hyperlink"):
		# 	inner_run = link.xpath("w:r", namespaces=link.nsmap)[0]
		# 	rId = link.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
		# 	print(inner_run.text, rId, paragraph.part.rels[rId]._target)
		if prun.bold and prun.italic:
			text += f"**_{ prun.text.strip() }_** "
		elif prun.bold:
			text += f"__{ prun.text.strip() }__ "
		elif prun.italic:
			text += f"_{ prun.text.strip() }_ "
		else:
			text += prun.text.strip() + " "
	return text.strip()


def convert_to_markdown(*paragraphs):
	md = ""
	start_of_list = True
	for paragraph in paragraphs:
		if paragraph.style.name == "List Paragraph":
			if start_of_list:
				md += "\n"
			md += "\n- " + convert_paragraph_text_to_markdown(paragraph)
			start_of_list = False
		else:
			md += "\n\n" + convert_paragraph_text_to_markdown(paragraph)
			start_of_list = True
	return md.strip()


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
					step.content = convert_to_markdown(*self.section[j + 1:i + 1])
				else:
					step.content = convert_to_markdown(*self.section[j:i + 1])
				self.animation.steps.append(step)
				j = i + 1
				i += 1
		else:
			self.animation.others[self.section[0].text] = convert_to_markdown(*self.section[1:])	
	
	def add_ressource_file(self, path):
		split = os.path.splitext(os.path.basename(path))
		self.animation.resources.append({
			"name": split[0],
			"ext": split[1],
			"size": os.path.getsize(path)
		})

	def parse_resources(self):
		for path in glob.glob(os.path.join(os.path.dirname(self.path), "*")):
			if path == self.path:
				continue
			if os.path.splitext(path)[1] == ".lnk":
				continue
			elif os.path.isfile(path):
				self.add_ressource_file(path)
			else:
				archive_path = path + ".zip"
				if os.path.isfile(archive_path):
					continue
				shutil.make_archive(path, "zip", path)
				self.add_ressource_file(archive_path)

	def parse(self, ressources=False):
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
		if ressources:
			self.parse_resources()


def parse_file(path, ressources=False):
	parser = DocumentParser(path)
	parser.parse(ressources=ressources)
	return parser.animation


def save_animation(database_path, animation_path, animation):
	database = {}
	if os.path.isfile(database_path):
		with open(database_path, "r", encoding="utf8") as file:
			database = json.load(file)
	database[os.path.realpath(animation_path)] = animation.to_dict()
	with open(database_path, "w", encoding="utf8") as file:
		json.dump(database, file, default=str, indent=4, sort_keys=True)


def main():
	parser = argparse.ArgumentParser()
	parser.add_argument("path", type=str, help="path to a DOCX file or a folder")
	parser.add_argument("-o", "--output", type=str, help="path to the output JSON file", default="db.json")
	parser.add_argument("-r", "--ressources", action="store_true", help="gather ressources")
	args = parser.parse_args()
	if os.path.isfile(args.path):
		paths = [args.path]
	else:
		path = args.path
		if path.endswith("\""):
			path = path[:-1]
		paths = []
		for root, dirs, files in os.walk(path, topdown=True):
			if "ATELIERS IDÉES" in root:
				continue
			for f in files:
				if f == "Fiche animation.docx":
					paths.append(os.path.join(root, f))					
	for animation_path in tqdm.tqdm(paths, unit="file"):
		animation = parse_file(animation_path, ressources=args.ressources)
		save_animation(args.output, animation_path, animation)


if __name__ == "__main__":
	main()
