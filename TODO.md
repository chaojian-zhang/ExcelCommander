# TODO

- (ExcelCommander) Separate dedicated Excel file writer and commander as two separate components.
Socket simplification: Just send a single string, and receiver always treat as multi-line scripts. No more payload concept. -> All we need to do is to Change server handler to deal with multiple incoming requests. Add frame size at sender, and the receiver don't otherwise need to do any special handling.
- (Naming) Call it Cell Script 😆

## Done
