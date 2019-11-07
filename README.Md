# xls export to some file tool  -- finish

## file about:
- Json
- Lua

## exls format:
### Files:

- cmd             column 2: params 
- CREATE          columns: first export file, column order by numbers.  
- USETYPE         columns: export limit of authority with 's' or 'c'   
- FLDTYPE         columns: use data type, with "Integer", "String", "Float"
- NOTE            columns: Field description.
- FLDNAME         columns: Field name.
- DATA            columns: Field data.
- CLOSE           columns: stop export data.

### forcus: if first column is empty or nothing, then export nothing this row.
