# Excel2SQL

Getting Excel data out of Excel files and into SQL Server is a bit of a pain.
Generally 'experts' refer you to SSIS, which is no good if the files you are receiving periodically regulary change formats, or if the files contain filtered tables, pivot dumps, and the like.

Here will appear some work in progress along two lines:
1. transforming Excel files to a set of delimited files (say CSV) without any loss of data or precision due to formatting, before getting these delimited files into SQL server using BULK INSERT or bcp.
2. for more 'standard' but still morphing Excel data, some SQL scripts to allow detecting the morphs, catering for them, and import in SQL server using OPENROWSET or linkd server approaches.
Both approaches are way faster than most SSIS approaches and (usually) far easier to maintain.

The problem is regularly discussed (without definitive solutions so far) on several forums, e.g.
http://www.sqlservercentral.com/Forums/Topic1664634-364-1.aspx

Please note that (amazingly) handling .CSV files seems not formally 'supported' by any MS tool (couldn't find the page that actually states this unambiguously, that fast, so 'seems' instead of 'is' for now).

Note: only since SQL2016 MS provides a FORMAT = 'CSV' switch in BULK INSERT making it treat CSVs like Excel does, thus solving some of the pain.
