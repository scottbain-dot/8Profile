# Synthetic test data

Everything in this folder is made up. Names are placeholders and every address
uses the reserved `example.edu` domain, so nothing here can match a real person
or a real FIS account.

- `students.csv`: the shape of the **Students** tab (Email · Name · Class). Paste into a
  test Sheet whose Config `domain` is set to `example.edu`, or just use the browser
  preview (`node dev/build-preview.js`), which seeds the same rows automatically.

Rules (see `PRIVACY.md`):
- Never put a real roster, real names, real scores or real photos anywhere in this repo.
- Real data lives only in the FIS-owned Google Sheet and is read at runtime by the
  signed-in student's own session.
