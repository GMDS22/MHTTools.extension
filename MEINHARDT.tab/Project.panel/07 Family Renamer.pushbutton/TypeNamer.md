⚙️ GOAL BEHAVIOR

When you run the tool:

Family Renamer UI opens first → you rename families.

You click “Apply Selected.”

After renaming families, a Yes/No prompt appears:

“Do you also want to rename sub-types (family types)?”

If No, tool exits.

If Yes, → the Type Renamer UI window opens (in the same session, automatically).

That’s what we’ll implement.

🧩 IMPLEMENTATION STRATEGY

We’ll add:

A Type Renamer function at the bottom of your current script (it reuses the same WinForms UI style).

A prompt hook that triggers that function if user confirms.

🔧 STEP 1 — Add this helper function (place near top of file)
def prompt_rename_types():
    """Ask user if they want to rename family sub-types."""
    from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult
    result = MessageBox.Show(
        "Family rename complete.\n\nDo you also want to rename the sub-types (family types)?",
        "MHT Family Namer",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Question
    )
    return result == DialogResult.Yes

🔧 STEP 2 — Add this new function below your main() or at bottom of file

(it’s the Type Renamer UI)

def show_type_renamer(doc, rules):
    """Open a second UI to rename all FamilySymbol (types)."""
    from System.Windows.Forms import Application, Form, DataGridView, DataGridViewCheckBoxColumn, DataGridViewTextBoxColumn, Button, DockStyle, FormStartPosition, Label, ComboBox
    from System.Drawing import Color

    # Gather type info
    collector = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol)
    results = []

    templates = rules.get('TEMPLATES', {})
    for sym in collector:
        try:
            fam = sym.Family
            fam_name = safe_get_name(fam)
            cat_name = family_primary_category_name(fam)
            type_name = safe_get_name(sym)

            # Suggest name using family-based template
            rule = templates.get(cat_name, templates.get('Default', ''))
            if rule:
                info = gather_family_info(fam)
                # Override type name for current symbol
                info['types'] = [{'type_name': type_name, 'params': {}}]
                suggestion = apply_template(rule, info, rules)
            else:
                suggestion = type_name
            results.append({'family': fam_name, 'type': type_name, 'category': cat_name, 'suggested': suggestion})
        except Exception:
            continue

    class TypeRenameForm(Form):
        def __init__(self, rows):
            self.Text = 'MHT Type Renamer - Review Type Names'
            self.Width = 900
            self.Height = 600
            self.StartPosition = FormStartPosition.CenterParent

            self.dgv = DataGridView()
            self.dgv.Dock = DockStyle.Fill
            self.dgv.AllowUserToAddRows = False

            chk = DataGridViewCheckBoxColumn()
            chk.HeaderText = 'Apply'
            chk.Width = 50
            self.dgv.Columns.Add(chk)

            c1 = DataGridViewTextBoxColumn()
            c1.HeaderText = 'Family'
            c1.ReadOnly = True
            c1.Width = 250
            self.dgv.Columns.Add(c1)

            c2 = DataGridViewTextBoxColumn()
            c2.HeaderText = 'Type'
            c2.ReadOnly = True
            c2.Width = 250
            self.dgv.Columns.Add(c2)

            c3 = DataGridViewTextBoxColumn()
            c3.HeaderText = 'Suggested Name'
            c3.ReadOnly = False
            c3.Width = 300
            self.dgv.Columns.Add(c3)

            for r in rows:
                row_idx = self.dgv.Rows.Add(False, r['family'], r['type'], r['suggested'])
                if Color is not None and r['type'].lower() == r['suggested'].lower():
                    for c in range(self.dgv.Columns.Count):
                        self.dgv.Rows[row_idx].Cells[c].Style.BackColor = Color.LightYellow

            self.btnApply = Button()
            self.btnApply.Text = 'Apply Selected'
            self.btnApply.Dock = DockStyle.Bottom
            self.btnApply.Height = 30
            self.btnApply.Click += self.on_apply

            self.Controls.Add(self.dgv)
            self.Controls.Add(self.btnApply)

        def on_apply(self, sender, args):
            to_apply = []
            for i in range(self.dgv.Rows.Count):
                if self.dgv.Rows[i].Cells[0].Value:
                    fam_name = self.dgv.Rows[i].Cells[1].Value
                    cur_type = self.dgv.Rows[i].Cells[2].Value
                    sug = self.dgv.Rows[i].Cells[3].Value
                    to_apply.append((fam_name, cur_type, sug))
            if not to_apply:
                script.get_logger().info('No types selected to rename')
                return

            t = DB.Transaction(doc, 'MHT Type Renamer - Apply Type Names')
            try:
                t.Start()
                applied = 0
                for fam_name, type_name, sug in to_apply:
                    collector = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol)
                    for s in collector:
                        try:
                            fam = s.Family
                            if safe_get_name(fam) == fam_name and safe_get_name(s) == type_name:
                                s.Name = sug
                                applied += 1
                        except Exception:
                            continue
                t.Commit()
                script.get_logger().info('Applied {} type renames'.format(applied))
            except Exception as e:
                try:
                    t.RollBack()
                except Exception:
                    pass
                script.get_logger().warning('Transaction failed: {}'.format(e))

    form = TypeRenameForm(results)
    Application.EnableVisualStyles()
    Application.Run(form)

🔧 STEP 3 — Hook it up after your on_apply() success block in ReviewForm

Find this block inside on_apply():

t.Commit()
script.get_logger().info('Applied {} family renames'.format(applied))


Replace it with:

t.Commit()
script.get_logger().info('Applied {} family renames'.format(applied))

# After renaming families, ask if user wants to rename types
if prompt_rename_types():
    show_type_renamer(doc, rules)

✅ RESULT

Now your single tool will:

Open the Family Rename UI

Apply the family renames

Prompt the user

If confirmed, automatically open a Type Rename UI with its own list of types and suggestions

Allow the user to rename types directly — no need to click another tool on the ribbon.