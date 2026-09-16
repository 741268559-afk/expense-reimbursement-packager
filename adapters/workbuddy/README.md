# WorkBuddy Adapter

WorkBuddy's public Skill format requires localized frontmatter fields that are not accepted by every strict Agent Skills validator. The repository therefore keeps a conservative universal `SKILL.md` and generates a separate WorkBuddy import package.

Build both packages:

```bash
python scripts/build_release_packages.py --out dist
```

Upload `expense-reimbursement-packager-workbuddy-v<version>.zip` in WorkBuddy through **Add Skill > Upload Skill**. The generated package contains:

- WorkBuddy-compatible localized `SKILL.md` frontmatter.
- `manifest.yaml` for enterprise Skill management.
- The same platform-neutral Python scripts and references as the universal package.
- The authorized built-in `费用报销单模板.xlsx` and its validated `form_cells.json` mapping.

No reimbursement data, approval screenshot, or user profile is included.
