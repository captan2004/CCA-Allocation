# CCA-Allocation
This code helps to allocate students to CCAs according to preferences and their rankings by the CCAs.

Functions:
1. Allocating DSA students to their respective CCAs
2. Allocating MEP students to their 1st choice Performing Arts CCA
3. Allocate students to either their 1st, 2nd or 3rd choice CCA if they have rankings
4. Allocate remaining students to any of their choice CCA if the CCA has vacancies for them
5. Allocate remaining students who still do not have CCAs randomly to CCAs that have vacancies

Explanation of Code:
First, we allocated the DSA students and MEP students. MEP students were given priority such that even if they do not have a rank or have a low rank, they still get their 1st choice Performing Arts CCA.
We chose to first allocate students to their top 3 choice CCA as long as they have a ranking within the CCA. However, since there is a limited quota, we decided to remove the extra students based on rank to benefit the CCA.
This results in some students getting allocated to a lower choice CCA because of their higher rank in the CCA or they were removed from the CCA they were first allocated to due to their lower rank.
Eg. 3d515955445d564c2e (Sample Allocation: Shooting, 1st choice) (Our algorithm: GE, 3rd choice)
This also results in some students getting allocated to a higher choice CCA.
Eg. 3d51595047565f4d28 (Sample Allocation: BB, 5th choice) (Our algorithm: Cricket, 2nd choice)

Then, we allocated the remaining students to any of their choice CCA if the CCA has vacancies for them through 9 rounds. 1st round: 1st choice, 2nd round: 2nd choice ...
This will result in some difference between the sample allocation as we want students to get their most preferred CCA. There will be some students who get allocated to a higher choice CCA as compared to the sample allocation.
Eg. 3d5159504357574c2a (Sample Allocation: NPCC, 3rd choice) (Our algorithm: RICO, 1st choice)
There will be students getting lower choice CCAs as compared to the sample allocation as spots have been taken up by others who rank the same CCA a higher choice.
Eg. 3d515950415956442a (Sample Allocation: RV, 6th choice) (Our algorithm: NPCC, 8th choice)
CCAs who have many people fighting for 1 spot will just take the first person in the dictionary as there is no way of differentiating who is more suited. Other students would then continue going through the rounds of allocation.

There will be however students that fail to be allocated to their top 9 choice CCAs if all the CCAs are full after all the rounds of allocation as students who have ranked that CCA higher would have gotten the place.
Eg. 3d5159574a5c5d452e (Sample Allocation: RIMB, 6th choice) (Our algorithm: RICO, No Choice)
Finally, they will be randomly allocated to CCAs that still have vacancies left. All students would then be allocated to a CCA.
