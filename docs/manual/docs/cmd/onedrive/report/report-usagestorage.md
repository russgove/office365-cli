# onedrive report usagestorage

Gets the trend on the amount of storage you are using in OneDrive for Business

## Usage

```sh
onedrive report usagestorage [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-p, --period <period>`|The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`
`-f, --outputFile [outputFile]`|Path to the file where the report should be stored in
`-o, --output [output]`|Output type. `text,json`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Gets the trend on the amount of storage you are using in OneDrive for Business for the last week

```sh
onedrive report usagestorage --period D7
```

Gets the trend on the amount of storage you are using in OneDrive for Business for the last week and exports the report data in the specified path in text format

```sh
onedrive report usagestorage --period D7 --output text --outputFile 'C:/report.txt'
```

Gets the trend on the amount of storage you are using in OneDrive for Business for the last week and exports the report data in the specified path in json format

```sh
onedrive report usagestorage --period D7 --output json --outputFile 'C:/report.json'
```
