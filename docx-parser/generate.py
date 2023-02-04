import argparse
import json
import os

import urllib.parse

from animation import Animation


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



def generate_animation(animation, folder):
	os.makedirs(folder, exist_ok=True)
	path = os.path.join(folder, animation.title + ".md")
	with open(path, "w", encoding="utf8") as file:
		file.write(f"# { animation.title }\n\n")
		if animation.subtitle is not None:
			file.write(f"{ animation.subtitle }\n\n")
		metadata_keys = {
			"topics": "Thématiques",
			"participants": "Participants",
			"duration": "Durée",
			"audience": "Public",
			"prerequisites": "Prérequis",
			"material": "Matériel"
		}
		for metadata_key, label in metadata_keys.items():
			value = getattr(animation.metadata, metadata_key)
			if value is None:
				continue
			file.write(f"**{ label }.**\n")
			if isinstance(value, list):
				for value_item in value:
					file.write(f"- { value_item }\n")
				file.write("\n")
			else:
				file.write(f"{ value }\n\n")
		for key, content in animation.others.items():
			file.write(f"## { key }\n\n{ content }\n\n")
		file.write("## Déroulé\n\n")
		for step in animation.steps:
			if step.title is not None:
				file.write(f"### { step.title } ({ step.duration } min)\n\n")
			file.write(f"{ step.content }\n\n")
		if animation.resources:
			file.write("## Ressources\n\n")
			for res in animation.resources:
				file.write(f"- [{ res['name' ]} ({ res['ext'][1:].upper() }, { sizeof_fmt(res['size']) })]({ urllib.parse.quote(res['name'] + res['ext']) })\n")


def main():
	parser = argparse.ArgumentParser()
	parser.add_argument("path", type=str, help="path to the JSON database file")
	parser.add_argument("-o", "--output", type=str, help="path to the output folder", default="generated")
	args = parser.parse_args()
	with open(args.path, "r", encoding="utf8") as file:
		data = json.load(file)
		for animation_dict in data.values():
			animation = Animation.from_dict(animation_dict)
			if animation.title is None:
				continue
			generate_animation(animation, args.output)



if __name__ == "__main__":
	main()