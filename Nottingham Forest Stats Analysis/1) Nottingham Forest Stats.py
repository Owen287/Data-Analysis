# Import packages
from bs4 import BeautifulSoup
import pandas as pd
import xlsxwriter

# Get the url and extract the necessary tables (1)
url1 = "https://fbref.com/en/squads/e4a775cb/2024-2025/c9/Nottingham-Forest-Stats-Premier-League"
tables_2024 = pd.read_html(url1)
url2 = "https://fbref.com/en/squads/e4a775cb/2023-2024/c9/Nottingham-Forest-Stats-Premier-League"
tables_2023 = pd.read_html(url2)
# After having located the table I want will extract the required data

# This is the table with fixtures - The teams xG and xGA
scores_2024 = tables_2024[1]
scores_2024 = scores_2024.dropna(subset = ["Result"]) # Removing games that havent been played yet
scores_2024 = scores_2024[["Venue", "Result", "GF", "GA", "xG", "xGA", "Poss"]] # Selecting important data
# Repeat for 2023
scores_2023 = tables_2023[1]
scores_2023 = scores_2023.dropna(subset = ["Result"]) # Removing games that havent been played yet
scores_2023 = scores_2023[["Venue", "Result", "GF", "GA", "xG", "xGA", "Poss"]] # Selecting important data

# Creating two datasets of Home and Away Games
h_2024 = scores_2024[scores_2024["Venue"] == "Home"]
a_2024 = scores_2024[scores_2024["Venue"] == "Away"]
# Repeat for 2023
h_2023 = scores_2023[scores_2023["Venue"] == "Home"]
a_2023 = scores_2023[scores_2023["Venue"] == "Away"]

# Making W/D/L chart with percentages
# Create a function to count the W/D/L and give a percentage of each - due to repeats
def WDL(x):
    win = x["Result"].str.count("W").sum()
    draw = x["Result"].str.count("D").sum()
    loss = x["Result"].str.count("L").sum()
    win_per = (win / (win + draw + loss)) * 100
    draw_per = (draw / (win + draw + loss)) * 100
    loss_per = (loss / (win + draw + loss)) * 100
    result_df = pd.DataFrame({"Result": ["W", "D", "L"],
                                "Count": [win, draw, loss],
                                "Percentage": [win_per, draw_per, loss_per],
                            })
    return result_df

# Creating a function to calculate average xG, xGA, GF, GA and Possession (Then separate for home and away)
def AvgMatchStats(x):
    avg_xG = x["xG"].mean()
    avg_GF = x["GF"].mean()
    avg_xGA = x["xGA"].mean()
    avg_GA = x["GA"].mean()
    avg_Possession = x["Poss"].mean()
    x_diff = avg_xG - avg_xGA
    A_diff = avg_GF - avg_GA
    total_xG = x["xG"].sum()
    total_GF = x["GF"].sum()
    xG_Diff = total_xG - total_GF
    result_df = pd.DataFrame({"Avg Expected Goals":[avg_xG], 
                              "Avg Goals For":[avg_GF],
                              "Expected Goal Discrepancy (xG - GF)":[xG_Diff], 
                              "Avg Expected Goals Against":[avg_xGA], 
                              "Avg Goals Against":[avg_GA],
                              "Expected Goal Difference (xG - xGA)":[x_diff],
                              "Actual Goal Difference":[A_diff],
                              "Average Possession":[avg_Possession]})
    return result_df

# Now want to start looking at individual players stats - Attacking players from this table
# Are players outperforming xG? and a quick look into xA
    # Will be doing everything per 90 in order to allow for more accurate comparison between this and
    # last season as we are not finished with this season
    
# Select the correct table
StandardStats_2024 = tables_2024[0]
StandardStats_2023 = tables_2023[0]
# Selecting Relevant Data
StandardStats_2024 = StandardStats_2024[[("Unnamed: 0_level_0", "Player"),
                                         ("Playing Time","90s"), ("Performance","Gls"), 
                                         ("Performance", "Ast"), ("Performance", "G+A"), 
                                         ("Performance", "G-PK"), ("Performance", "PK"),
                                         ("Expected", "xG"), ("Expected", "npxG"), ("Expected", "xAG"),
                                         ("Expected", "npxG+xAG"), ("Progression", "PrgC"),
                                         ("Progression", "PrgP"), ("Progression", "PrgR")]]
StandardStats_2023 = StandardStats_2023[[("Unnamed: 0_level_0", "Player"),
                                         ("Playing Time","90s"), ("Performance","Gls"), 
                                         ("Performance", "Ast"), ("Performance", "G+A"), 
                                         ("Performance", "G-PK"), ("Performance", "PK"),
                                         ("Expected", "xG"), ("Expected", "npxG"), ("Expected", "xAG"),
                                         ("Expected", "npxG+xAG"), ("Progression", "PrgC"),
                                         ("Progression", "PrgP"), ("Progression", "PrgR")]]
# Removing top columns
StandardStats_2024.columns = StandardStats_2024.columns.droplevel(0)
StandardStats_2023.columns = StandardStats_2023.columns.droplevel(0)
# Function to extract key stats
def StandardStatsPer90(x):
    # Preserve the player and 90s
    Players = x["Player"]
    AmountOf90s = x["90s"]
    #Drop Player and 90s column
    x = x.drop(columns = ["Player", "90s"])
    # Divide every stat by the amount of 90s played
    x = x.div(AmountOf90s, axis = 0)
    return x
# Maybe concat the stats and the stats per 90 after?

# In Depth Defensive Stats
DefensiveStats_2024 = tables_2024[8]
DefensiveStats_2023 = tables_2023[8]
# Removing Nation, Age and Matchhes
DefensiveStats_2024 = DefensiveStats_2024.drop(columns = [("Unnamed: 1_level_0", "Nation"), 
                                                          ("Unnamed: 3_level_0", "Age"), 
                                                          ("Unnamed: 21_level_0", "Matches")])
DefensiveStats_2023 = DefensiveStats_2023.drop(columns = [("Unnamed: 1_level_0", "Nation"), 
                                                          ("Unnamed: 3_level_0", "Age"), 
                                                          ("Unnamed: 21_level_0", "Matches")])
# Create a function to generate defensive actions per 90
def DefensiveStatsPer90(x):
    # Preserve Stats that cant be divided by 90
    Preserved = [("Unnamed: 0_level_0", "Player"), ("Unnamed: 2_level_0", "Pos"), ("Unnamed: 4_level_0", "90s"),
                  ("Challenges", "Tkl%")]
    # Acquire 90s played for each player
    AmountOf90s = x["Unnamed: 4_level_0", "90s"]
    # Drop stats that can't be divided
    x = x.drop(columns = Preserved)
    # Divide remaining stats by 90s
    x = x.div(AmountOf90s, axis = 0)
    return x

# Goal and Shot Creation Stats
ShotGoalCreation_2024 = tables_2024[7]
ShotGoalCreation_2023 = tables_2023[7]
# Selecting Relevant Columns
ShotGoalCreation_2024 = ShotGoalCreation_2024[[("Unnamed: 0_level_0", "Player"), ("Unnamed: 2_level_0", "Pos"),
                                              ("Unnamed: 4_level_0", "90s"), ("SCA", "SCA"), ("SCA", "SCA90"),
                                              ("GCA", "GCA"), ("GCA", "GCA90")]]
ShotGoalCreation_2023 = ShotGoalCreation_2023[[("Unnamed: 0_level_0", "Player"), ("Unnamed: 2_level_0", "Pos"),
                                              ("Unnamed: 4_level_0", "90s"), ("SCA", "SCA"), ("SCA", "SCA90"),
                                              ("GCA", "GCA"), ("GCA", "GCA90")]]

# Possession Stats - Progressive Carries, Take-Ons, Recieving Progressive Passes (Per 90)
PossessionStats_2024 = tables_2024[9]
PossessionStats_2023 = tables_2023[9]
# Selecting Relevant Columns 
PossessionStats_2024 = PossessionStats_2024[[("Unnamed: 0_level_0", "Player"), ("Unnamed: 2_level_0", "Pos"),
                                            ("Unnamed: 4_level_0", "90s"), ("Take-Ons", "Att"), ("Take-Ons", "Succ"),
                                            ("Take-Ons", "Succ%"), ("Take-Ons", "Tkld"), ("Take-Ons", "Tkld%"),
                                            ("Carries", "Carries"), ("Carries", "TotDist"), ("Carries", "PrgDist"),
                                            ("Carries", "PrgC"), ("Carries", "1/3"), ("Carries", "CPA"), 
                                            ("Carries", "Mis"), ("Carries", "Dis")]]
PossessionStats_2023 = PossessionStats_2023[[("Unnamed: 0_level_0", "Player"), ("Unnamed: 2_level_0", "Pos"),
                                            ("Unnamed: 4_level_0", "90s"), ("Take-Ons", "Att"), ("Take-Ons", "Succ"),
                                            ("Take-Ons", "Succ%"), ("Take-Ons", "Tkld"), ("Take-Ons", "Tkld%"),
                                            ("Carries", "Carries"), ("Carries", "TotDist"), ("Carries", "PrgDist"),
                                            ("Carries", "PrgC"), ("Carries", "1/3"), ("Carries", "CPA"), 
                                            ("Carries", "Mis"), ("Carries", "Dis")]]
# Create a function to divide all of the stats by the amount of 90s
def PossessionStatsPer90(x):
    # Preserve Undividable Stats
    Preserved = [("Unnamed: 0_level_0", "Player"), ("Unnamed: 2_level_0", "Pos"), ("Unnamed: 4_level_0", "90s"),
                ("Take-Ons", "Succ%"), ("Take-Ons", "Tkld%")]
    # Extract the 90s played for each player
    AmountOf90s = x[("Unnamed: 4_level_0", "90s")]
    # Drop Preserved Columns
    x = x.drop(columns = Preserved)
    # Divide each stat by amount of 90s
    x = x.div(AmountOf90s, axis = 0)
    return x

# Going to look at Passing Stats
Passing_Stats_2024 = tables_2024[5]
Passing_Stats_2023 = tables_2023[5]
# Selecting Relevant Columns
Passing_Stats_2024 = Passing_Stats_2024[[("Unnamed: 0_level_0", "Player"), ("Unnamed: 2_level_0", "Pos"), 
                                     ("Unnamed: 4_level_0", "90s"), ("Total", "Cmp%"), ("Total", "TotDist"),
                                     ("Total", "PrgDist"), ("Unnamed: 27_level_0", "PrgP")]]
Passing_Stats_2023 = Passing_Stats_2023[[("Unnamed: 0_level_0", "Player"), ("Unnamed: 2_level_0", "Pos"), 
                                     ("Unnamed: 4_level_0", "90s"), ("Total", "Cmp%"), ("Total", "TotDist"),
                                     ("Total", "PrgDist"), ("Unnamed: 27_level_0", "PrgP")]]
# Removing top columns as arent necessary and makes the division function easier
Passing_Stats_2024.columns = Passing_Stats_2024.columns.droplevel(0)
Passing_Stats_2023.columns = Passing_Stats_2023.columns.droplevel(0)
# Create Function to calculate passing per 90 stats. Also it will look at percentage of distance of progressive
# passes to total distance. Will show who is more attacking focused and could explain low pass % completions
def PassingPer90(x):
    Preserved = ["Player", "Pos", "90s", "Cmp%"]
    AmountOf90s = x["90s"]
    ProgressiveDistancePercentage = (x["PrgDist"] / x["TotDist"]) * 100
    x = x.drop(columns = Preserved)
    x = x.div(AmountOf90s, axis = 0)
    x["PrgDist%"] = ProgressiveDistancePercentage
    return x
 
# Now for the final Misc Stats
MiscStats_2024 = tables_2024[11]
MiscStats_2023 = tables_2023[11]
# Selecting Relvant Columns
MiscStats_2024 = MiscStats_2024[[("Unnamed: 0_level_0", "Player"), ("Unnamed: 2_level_0", "Pos"),
                                 ("Unnamed: 4_level_0", "90s"), ("Performance", "Crs"),
                                 ("Performance", "Int"), ("Performance", "Recov"), ("Aerial Duels", "Won"),
                                 ("Aerial Duels", "Lost"), ("Aerial Duels", "Won%")]]
MiscStats_2023 = MiscStats_2023[[("Unnamed: 0_level_0", "Player"), ("Unnamed: 2_level_0", "Pos"),
                                 ("Unnamed: 4_level_0", "90s"), ("Performance", "Crs"),
                                 ("Performance", "Int"), ("Performance", "Recov"), ("Aerial Duels", "Won"),
                                 ("Aerial Duels", "Lost"), ("Aerial Duels", "Won%")]]
# Removing top columns
MiscStats_2024.columns = MiscStats_2024.columns.droplevel(0)
MiscStats_2023.columns = MiscStats_2023.columns.droplevel(0)
# Create a function to give Stats per 90
def MiscStatsPer90(x):
    Preserved = ["Player", "Pos", "90s", "Won%"]
    AmountOf90s = x["90s"]
    x = x.drop(columns = Preserved)
    x = x.div(AmountOf90s, axis = 0)
    return x

# Now to run all the tables through the functions and concat the per 90 stats to the right of the normal stats
# Fixtures 

# Also want the raw data for the Scores and Fixtures tables. Potential for statistical comparison with the raw data
# any significant difference in certain stats for this and last season?
raw_2024 = tables_2024[1]
raw_2023 = tables_2023[1]
raw_cooper_2023 = raw_2023.iloc[0:17]
raw_post_cooper_2023 = raw_2023.iloc[17:]
# Overall Fixtures (Need to split into before and after cooper sacking)
Results_2024 = WDL(scores_2024)
Results_2023 = WDL(scores_2023)
# Home Results
Home_Results_2024 = WDL(h_2024)
Away_Results_2024 = WDL(a_2024)
# Away Results
Home_Results_2023 = WDL(h_2023)
Away_Results_2023 = WDL(a_2023)
# 2023 Season - Before and After Cooper Sacking
Cooper2023 = scores_2023.iloc[0:17]
Post_Cooper2023 = scores_2023.iloc[17:]

Cooper_Home_2023 = Cooper2023[Cooper2023["Venue"] == "Home"]
Cooper_Away_2023 = Cooper2023[Cooper2023["Venue"] == "Away"]
Cooper_Results_2023 = WDL(Cooper2023)
Cooper_Home_Results_2023 = WDL(Cooper_Home_2023)
Cooper_Away_Results_2023 = WDL(Cooper_Away_2023)

Post_Cooper_Home_2023 = Post_Cooper2023[Post_Cooper2023["Venue"] == "Home"]
Post_Cooper_Away_2023 = Post_Cooper2023[Post_Cooper2023["Venue"] == "Away"]
Post_Cooper_Results_2023 = WDL(Post_Cooper2023)
Post_Cooper_Home_Results_2023 = WDL(Post_Cooper_Home_2023)
Post_Cooper_Away_Results_2023 = WDL(Post_Cooper_Away_2023)

# Repeating this but looking at match stats 
Team_Stats_2024 = AvgMatchStats(scores_2024)
Team_Stats_2023 = AvgMatchStats(scores_2023)

Team_Home_Stats_2024 = AvgMatchStats(h_2024)
Team_Away_Stats_2024 = AvgMatchStats(a_2024)

Team_Home_Stats_2023 = AvgMatchStats(h_2023)
Team_Away_Stats_2023 = AvgMatchStats(a_2023)

Cooper_Team_Stats_2023 = AvgMatchStats(Cooper2023)
Cooper_Team_Stats_Home_2023 = AvgMatchStats(Cooper_Home_2023)
Cooper_Team_Stats_Away_2023 = AvgMatchStats(Cooper_Away_2023)
Post_Cooper_Team_Stats_2023 = AvgMatchStats(Post_Cooper2023)
Post_Cooper_Team_Stats_Home_2023 = AvgMatchStats(Post_Cooper_Home_2023)
Post_Cooper_Team_Stats_Away_2023 = AvgMatchStats(Post_Cooper_Away_2023)


# Players Stats
# StandardStats
Overall_Player_Stats_Per90_2024 = StandardStatsPer90(StandardStats_2024)
Overall_Player_Stats_Per90_2023 = StandardStatsPer90(StandardStats_2023)

# DefensiveStats
Defensive_Stats_Per90_2024 = DefensiveStatsPer90(DefensiveStats_2024)
Defensive_Stats_Per90_2023 = DefensiveStatsPer90(DefensiveStats_2023)

# ShotGoalCreation

# PossessionStats
Possession_Stats_Per90_2024 = PossessionStatsPer90(PossessionStats_2024)
Possession_Stats_Per90_2023 = PossessionStatsPer90(PossessionStats_2023)

# Passing
Passing_Stats_Per90_2024 = PassingPer90(Passing_Stats_2024)
Passing_Stats_Per90_2023 = PassingPer90(Passing_Stats_2023)

# MiscStats
Misc_Stats_Per90_2024 = MiscStatsPer90(MiscStats_2024)
Misc_Stats_Per90_2023 = MiscStatsPer90(MiscStats_2023)


# Concat all the dataframes together
# Team Data
Final_Results_2024 = pd.concat([Results_2024, Home_Results_2024, Away_Results_2024], axis = 1)
Final_Results_2023 = pd.concat([Results_2023, Home_Results_2023, Away_Results_2023], axis = 1)
Final_Cooper_Results = pd.concat([Cooper_Results_2023, Cooper_Home_Results_2023, Cooper_Away_Results_2023], axis = 1)
Final_Post_Cooper_Results = pd.concat([Post_Cooper_Results_2023, Post_Cooper_Home_Results_2023, 
                                       Post_Cooper_Away_Results_2023], axis = 1)
Final_Team_Stats_2024 = pd.concat([Team_Stats_2024, Team_Home_Stats_2024, Team_Away_Stats_2024], axis = 1)
Final_Team_Stats_2023 = pd.concat([Team_Stats_2023, Team_Home_Stats_2023, Team_Away_Stats_2023], axis = 1)
Final_Cooper_Stats = pd.concat([Cooper_Team_Stats_2023, Cooper_Team_Stats_Home_2023, 
                                Cooper_Team_Stats_Away_2023], axis = 1)
Final_Post_Cooper_Stats = pd.concat([Post_Cooper_Team_Stats_2023, Post_Cooper_Team_Stats_Home_2023,
                                     Post_Cooper_Team_Stats_Away_2023], axis = 1)

# Player Data
Overall_Player_Stats_2024 = pd.concat([StandardStats_2024, Overall_Player_Stats_Per90_2024], axis = 1)
Overall_Player_Stats_2023 = pd.concat([StandardStats_2023, Overall_Player_Stats_Per90_2023], axis = 1)

Defence_2024 = pd.concat([DefensiveStats_2024, Defensive_Stats_Per90_2024], axis = 1)
Defence_2023 = pd.concat([DefensiveStats_2023, Defensive_Stats_Per90_2023], axis = 1)

Possession_2024 = pd.concat([PossessionStats_2024, Possession_Stats_Per90_2024], axis = 1)
Possession_2023 = pd.concat([PossessionStats_2023, Possession_Stats_Per90_2023], axis = 1)

Passing_2024 = pd.concat([Passing_Stats_2024, Passing_Stats_Per90_2024], axis = 1)
Passing_2023 = pd.concat([Passing_Stats_2023, Passing_Stats_Per90_2023], axis = 1)

Misc_2024 = pd.concat([MiscStats_2024, Misc_Stats_Per90_2024], axis = 1)
Misc_2023 = pd.concat([MiscStats_2023, Misc_Stats_Per90_2023], axis = 1)



# Use xlxs writer to insert titles for my dataframes. Then using ExcelWriter to insert the data frames
with pd.ExcelWriter("Nottingham Forest Stats.xlsx", engine = "xlsxwriter") as writer:
    # Team Stats
    Final_Results_2024.to_excel(writer, sheet_name="Team Data", startrow = 2, startcol = 0, index = False)
    Final_Results_2023.to_excel(writer, sheet_name="Team Data", startrow=9, startcol=0, index = False)
    Final_Cooper_Results.to_excel(writer, sheet_name="Team Data", startrow=16, startcol=0, index=False)
    Final_Post_Cooper_Results.to_excel(writer, sheet_name="Team Data", startrow=23, startcol=0, index=False)
    Final_Team_Stats_2024.to_excel(writer, sheet_name="Team Data", startrow=2, startcol=10, index=False)
    Final_Team_Stats_2023.to_excel(writer, sheet_name="Team Data", startrow=9, startcol=10, index=False)
    Final_Cooper_Stats.to_excel(writer, sheet_name="Team Data", startrow=16, startcol=10, index=False)
    Final_Post_Cooper_Stats.to_excel(writer, sheet_name="Team Data", startrow=23, startcol=10, index=False)

    # Creating the Worksheet
    results_data = writer.sheets["Team Data"]
    # List of all the additional Titles I want to add
    Team_Data_Titles = [("A1", "Results 2024"), ("A2", "Combined Results"), ("D2", "Home Results"), ("G2", "Away Results"),
                    ("A8", "Results 2023"), ("A9", "Combined Results"), ("D9", "Home Results"), ("G9", "Away Results"),
                    ("A15", "Cooper Results 2023"), ("A16", "Combined Results"), ("D16", "Home Results"), ("G16", "Away Results"),
                    ("A22", "Post Cooper / Nuno Results 2023"), ("A23", "Combined Results"), ("D23", "Home Results"), ("G23", "Away Results"),
                    ("K1", "Team Stats 2024"), ("K2", "Combined Stats"), ("S2", "Home Stats"), ("AA2", "Away Stats"),
                    ("K8", "Team Stats 2023"), ("K9", "Combined Stats"), ("S9", "Home Stats"), ("AA9", "Away Stats"),
                    ("K15", "Cooper Stats 2023"), ("K16", "Combined Stats"), ("S16", "Home Stats"), ("AA16", "Away Stats"),
                    ("K22", "Post Cooper / Nuno Stats 2023"), ("K23", "Combined Results"), ("S23", "Home Stats"), ("AA23", "Away Stats")
                    ] 
    # Loop through the titles to insert them where I want them
    for cell, title in Team_Data_Titles:
        results_data.write(cell, title)

    # Overall Player Stats
    Overall_Player_Stats_2024.to_excel(writer, sheet_name="Overall Player Stats", startrow=3, startcol=0, index=False)
    Overall_Player_Stats_2023.to_excel(writer, sheet_name="Overall Player Stats", startrow=35, startcol=0, index=False)
    # Creating the worksheet
    overall_data = writer.sheets["Overall Player Stats"]
    # Adding titles to the data
    Overall_Stats_Titles = [("A1", "Players Overall Stats"), ("A2", "Players Overall Stats 2024"), 
                        ("A3", "Total Stats"), ("O3", "Per 90 Stats"),
                        ("A34", "Players Overall Stats 2023"), ("A35", "Total Stats"), ("O35", "Per 90 Stats")
                        ]
    for cell, title in Overall_Stats_Titles:
        overall_data.write(cell, title)
    
    # ShotGoalCreation Stats
    # Inserting dataframes (have to use index=True as using MultiIndex columns)
    # (Also why the titles are shifted to the right one tile)
    ShotGoalCreation_2024.to_excel(writer, sheet_name="Shot and Goal Creation", startrow=2, startcol=0)
    ShotGoalCreation_2023.to_excel(writer, sheet_name="Shot and Goal Creation", startrow=33, startcol=0)
    # Creating the worksheet
    shotgoal = writer.sheets["Shot and Goal Creation"]
    # Add titles
    ShotGoalCreation_Titles = [("B1", "Shot and Goal Creation Stats"), 
                               ("B2", "Shot and Goal Creation Stats 2024"),
                               ("B33", "Shot and Goal Creation Stats 2023")
                               ]
    for cell, title in ShotGoalCreation_Titles:
        shotgoal.write(cell, title)
    
    # Defensive Stats
    Defence_2024.to_excel(writer, sheet_name="Defensive Stats", startrow=3, startcol=0)
    Defence_2023.to_excel(writer, sheet_name="Defensive Stats", startrow=34, startcol=0)
    # Creating the worksheet
    defence_data = writer.sheets["Defensive Stats"]
    # Adding titles to the data
    Defence_Titles = [("B1", "Defensive Stats"), 
                  ("B2", "Defensive Stats 2024"), ("B3", "Total Defensive Stats"), ("U3", "Defensive Stats Per 90"),
                  ("B33", "Defensive Stats 2023"), ("B34", "Total Defensive Stats"), ("U34", "Defensive Stats Per 90")
                  ]
    for cell, title in Defence_Titles:
        defence_data.write(cell, title)

    # Possession Stats
    Possession_2024.to_excel(writer, sheet_name="Possession Stats", startrow=3, startcol=0)
    Possession_2023.to_excel(writer, sheet_name="Possession Stats", startrow=35, startcol=0)
    # Creating worksheet
    possession_data = writer.sheets["Possession Stats"]
    # Adding titles to the data
    Possession_Titles = [("B1", "Possession Data"),
                     ("B2", "Possession Stats 2024"), ("B3", "Total Possession Stats"), ("R3", "Possession Stats Per 90"),
                     ("B34", "Possession Stats 2023"), ("B35", "Total Possession Stats"), ("R35", "Possession Stats Per 90")
                     ]
    for cell, title in Possession_Titles:
        possession_data.write(cell, title)
    
    # Passing Stats
    Passing_2024.to_excel(writer, sheet_name="Passing Stats", startrow=3, startcol=0, index=False)
    Passing_2023.to_excel(writer, sheet_name="Passing Stats", startrow=32, startcol=0, index=False)
    # Creating the worksheet
    passing_data = writer.sheets["Passing Stats"]
    # Adding titles
    Passing_Titles = [("A1", "Passing Stats"),
                  ("A2", "Passing Stats 2024"), ("A3", "Total Passing Stats"), ("H3", "Passing Per 90 Stats"),
                  ("A31", "Passing Stats 2023"), ("A32", "Total Passing Stats"), ("H32", "Passing Per 90 Stats")
                  ]
    for cell, title in Passing_Titles:
        passing_data.write(cell, title)
    
    # Misc Stats
    Misc_2024.to_excel(writer, sheet_name="Misc Stats", startrow=3, startcol=0)
    Misc_2023.to_excel(writer, sheet_name="Misc Stats", startrow=32, startcol=0)
    # Create the worksheet
    misc_data = writer.sheets["Misc Stats"]
    # Add titles
    Misc_Titles = [("A1", "Misc Stats"),
               ("A2", "Misc Stats 2024"), ("A3", "Total Misc Stats"), ("J3", "Misc Stats Per 90"),
               ("A31", "Misc Stats 2023"), ("A32", "Total Misc Stats"), ("J32", "Misc Stats Per 90")
               ]
    for cell, title in Misc_Titles:
        misc_data.write(cell, title)
    
    # Sheet with just the raw results data
    raw_2024.to_excel(writer, sheet_name="Raw Fixture Data", startrow=2, startcol=0, index=False)
    raw_2023.to_excel(writer, sheet_name="Raw Fixture Data", startrow=44, startcol=0, index=False)
    raw_cooper_2023.to_excel(writer, sheet_name="Raw Fixture Data", startrow=86, startcol=0, index=False)
    raw_post_cooper_2023.to_excel(writer, sheet_name="Raw Fixture Data", startrow=107, startcol=0, index=False)
    # Create the worksheet
    raw_data = writer.sheets["Raw Fixture Data"]
    # Add titles
    Raw_Titles = [("A1", "Raw Fixture Data"), ("A2", "Raw Fixture Data 2024"), ("A44", "Raw Fixture Data 2023"), 
                  ("A86", "Raw Cooper Fixture Data 2023"), ("A107", "Raw Post Cooper / Nuno Fixture Data 2023")]
    for cell, title in Raw_Titles:
        raw_data.write(cell, title)





    
    


    
    


    