# honeydoc
Insert tracking information into docx files

![HoneyDoc CLI](/images/honeydoc_cli.png)

## Notes
- Relationship Id="rId” supports strings.
- Default to Incrementing the largest existing ID, but have option to specify a custom string.
- Must ensure the content_types xml includes the imported content type (PNG in example).
- We can safely ignore the /word/setting.xml file.

## Example

![HoneyDoc Example](/images/example_beacon.png)