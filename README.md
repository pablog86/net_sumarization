# net_sumarization

Create an excel document with all the subnets to summarize in the first column (A) of the document using CIDR format (X.X.X.X/Y). 
This script will ask for the file and it will iterate through the list of values trying to find contiguous networks to agrupate in order to avoid un used gaps in the summarization.

The result will be printed in the console and writed in the same excel file using the column (B)
