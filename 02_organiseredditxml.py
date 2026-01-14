import xml.etree.ElementTree as ET
from collections import defaultdict

# Parse the existing XML file
input_file = 'PATH_TO_YOUR_INPUT_FILE.xml'
tree = ET.parse(input_file)
root = tree.getroot()

# Create a dictionary to hold posts by subreddit
subreddit_dict = defaultdict(list)

# Iterate through each row and organize data by subreddit
for row in root.findall('row'):
    subreddit = row.find('Subreddit').text
    post_id = row.find('PostID').text
    post_score = row.find('PostScore').text
    body = row.find('Body').text
    title = row.find('Title').text  # Pull the <title> text

    post_element = ET.Element('post', ID=post_id, PostScore=post_score)
    title_element = ET.SubElement(post_element, 'title')
    title_element.text = title
    body_element = ET.SubElement(post_element, 'body')
    body_element.text = body

    subreddit_dict[subreddit].append(post_element)

# Create the new root element
new_root = ET.Element('root')

# Add subreddits and their posts to the new root element
for subreddit, posts in subreddit_dict.items():
    subreddit_element = ET.Element('subreddit', name=subreddit)
    subreddit_element.extend(posts)
    new_root.append(subreddit_element)

# Write the new XML structure to a file
output_file = 'PATH_TO_YOUR_OUTPUT_FILE.xml'
new_tree = ET.ElementTree(new_root)
new_tree.write(output_file, encoding='utf-8', xml_declaration=True)