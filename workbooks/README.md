# workbooks/

Suggested local Excel drop folder.

Put your .xlsx files here. When you tell the agent in the terminal:

"open the local folder" / "yerel klasörü aç"

it opens this folder **even if you don't pass a path**. You can still
open any other path via `open_local_folder("F:\\other\\path")` — this
folder is just a tidy default, not a hard constraint.
