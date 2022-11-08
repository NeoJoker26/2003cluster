# Greens On Screen!

## COMP2003 Group Project

- Morgan Trainor – Client Liaison
- Jonathan Boxall – Product Owner
- Jacob Brimacombe – Scrum Master
- Kieran Kirkwood – Technical Lead



## How to use this github repo setup!

Firstly, no one should be pushing to the main branch what-so-ever. The technical lead will merge branches when it is safe to do so, or when milestones are completed.

Depending on the work being handled you will be pushing to Features, Bugs or Staging. The main branch reflects a state with the latest delivered development changes for the next release/stable version.

## Support Branches
These branches will be used to aid parallel development between the group which makes to track features, and to assist in quickly fixing live-production issues. These branches will have a limited life-span.

### Feature-name > Merge from main and push to staging. This is used to isolate a developing feature from the stable branch.

### Commands

$ git checkout -b feature-name main <creates a local branch for the new feature> 

$ git push origin feature-name <makes the new feature remotely available>

$ git merge main <merges changes from main into feature branch>

### When development of the feature has been completed, push to staging = never to main.

$ git checkout -b staging <checkout allows you to navigate branches>

$ git add yourfile.name <adding your selected files to push>

$ git commit -m "commit message here"

$ git push origin staging

This also reflects the same stages for the other supporting branches, which include hot-fixes and a bug branch.
