[filter1]: ## "Has word: intern OR recruit OR interview OR applying"
[filter2]: ## "(intern OR recruit OR interview OR applying) -unsubscribe -newsletter -promotion -reward -sale -credit -loan -statement -billing -receipt -points -offers)"
# Journal

### Entry 1
#### Apr 6, 2025
###### Creating filter is still an issue. 
- [Typical filter][filter1] is decent at accurately receiving emails and filtering correctly but struggles with job notifications and updates from linkedin, current solution is just to unsubscribe from those notifications. 
- Tried using a different [filter][filter2]  yesterday which increased my email count from ~260 to over 600, where more than 300 were incorrect. 
    - Need help filtering out college emails without iusing hardcoded words to exclude. Some internship applications include the term `On Campus` or such and would completely ignore them.
##### Fuzzy Search Alternatives
- Currently trying to figure out whether I should continue using fuzzy search to find similar jobs
    - Fuzzy search works great so far for determining companies without duplicates (6% Rate Error)
- However, fuzzy search is sometimes overlapping and removing applications so instead of having one company and two different job positions (two different rows), it will only make one row and update the job role
- I just realized that I'm not even doing fuzzy search on the roles, so I just added that, I need to test it out. I also added a build entry function as to not re use code.
    - I think I'm going to need to add some cleaning for the roles because as it stands the role names are VERY similar when adding entries even for obvious examples as in 2025 IT Intern vs. 2025 IT Intern – 31090529
- The results are much better however I'm going to need to add some arguements for when the job position cannot be extracted
    - The issue with fixing this and just allocating it to the nearest email with the same job position is that we could update the wrong email
    - Despite prompting model to leave any email that the position isn't specified  as N/A it is still outputting other answers, might be that my prompt is becoming to long which is entirely likely.



