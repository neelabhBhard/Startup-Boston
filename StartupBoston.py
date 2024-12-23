import pandas as pd

df = pd.read_excel ( "/Users/neelabhbhardwaj/Desktop/Startup Boston/SBW2024 All Attendee Data.xlsx")

df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

df['floor_num'] = df.apply(
    lambda row: 1 if row['1st_floor'] == 1 else (4 if row['4th_floor'] == 1 else (5 if row['5th_floor'] == 1 else None)),
    axis=1
)
num_of_att_each_session = (
    df.groupby(['Session', 'floor_num'])['email']
    .nunique()
    .reset_index(name='unique_email_count')
)

unique_count = df['email'].nunique()
print(f"Number of unique emails: {unique_count}")


# Display the result
print(num_of_att_each_session)


unique_companies = df[["company_name"]].drop_duplicates().reset_index(drop = True)
print(unique_companies)


frequency_by_continent = (
    df.groupby([ 'currently_reside'])['email']
    .nunique()
    .reset_index(name='unique_email_count')
)

print(frequency_by_continent)

frequency_by_job_role = (
    df.groupby([ 'current_job_function'])['email']
    .nunique()
    .reset_index(name='unique_email_count')
)

print(frequency_by_job_role)


frequency_by_state = (
    df.groupby([ 'state_you_currently_reside'])['email']
    .nunique()
    .reset_index(name='unique_email_count')
)

print(frequency_by_state)


frequency_by_company_size = (
    df.groupby([ 'current_company_size'])['email']
    .nunique()
    .reset_index(name='unique_email_count')
)

print(frequency_by_company_size)


frequency_by_funding_stage = (
    df.groupby([ 'funding_stage_of_your_startup'])['email']
    .nunique()
    .reset_index(name='unique_email_count')
)

print(frequency_by_funding_stage)




breakdown_df = df[["email", "Session","currently_reside", "state_you_currently_reside", "job_title","current_job_function","current_company_size", "funding_stage_of_your_startup"]]\
    .drop_duplicates(keep='first')\
    .dropna(how='all')\
    .reset_index(drop = True)
print(breakdown_df)

def get_unique_cols(col_name):
    lst = [[y.strip() for y in str(x).replace("[", "").replace("]", "").replace("'", "").split(',')] for x in list(df[col_name])]
    out_lst = list(set(item for sublist in lst for item in sublist))
    out_lst.remove("")
    out_lst.remove("nan")
    print(out_lst)
    print(len(out_lst))

def get_freq_of_questions(col_name):
    tmp_df = df.groupby('email')[col_name].agg(lambda x: ', '.join([str(i).strip("[]'") for i in x])).reset_index()
    print(tmp_df)
    lst = [list(set([y.strip() for y in str(x).replace("[", "").replace("]", "").replace("'", "").split(',')])) for x in list(tmp_df[col_name])]
    d = dict()
    for x in lst:
        for y in x:
            try:
                d[y] += 1
            except:
                d[y] = 1
    d.pop("")
    d.pop("nan")
    return pd.DataFrame(list(d.items()), columns=['Question', 'Frequency'])


def group_count_2d_col(cols, col_to_process):
    tmp_df = df[cols]\
        .drop_duplicates(keep='first')\
        .dropna(how='all')\
        .reset_index(drop = True)
    lst = [[y.strip() for y in str(x).replace("[", "").replace("]", "").replace("'", "").split(',')] for x in list(tmp_df[col_to_process])]
    d = dict()
    for x in lst:
        for y in x:
            try:
                d[y] += 1
            except:
                d[y] = 1
    d.pop("")
    d.pop("nan")
    return pd.DataFrame(list(d.items()), columns=['Question', 'Frequency'])


how_did_you_hear_about_this_event = get_freq_of_questions("how_did_you_hear_about_this_event")
why_are_you_attending_sbw2024 = get_freq_of_questions("why_are_you_attending_sbw2024")

# If you have dataframes in a dictionary
dfs = {
    'Attendance Summary': num_of_att_each_session,
    'Breakdown': breakdown_df,
    'How did you hear': how_did_you_hear_about_this_event,
    'Why are you attending': why_are_you_attending_sbw2024,
    'Job role frequency': frequency_by_job_role,
    'Company Size frequency': frequency_by_company_size,
    'Funding Stage Frequency': frequency_by_funding_stage,
    'Sate You currently live in': frequency_by_state,
    

}

with pd.ExcelWriter('/Users/neelabhbhardwaj/Desktop/bostonStartUp_attendance.xlsx') as writer:
    for sheet_name, tmp_df in dfs.items():
        tmp_df.to_excel(writer, sheet_name=sheet_name, index=False)


duplicate_mask = breakdown_df.duplicated()
print("Number of duplicate rows:", duplicate_mask.sum())

# df = df[["email", "first_name", "last_name", "1st_floor", "4th_floor", "5th_floor"]]
#
# print(df)
# unique_rows = df.drop_duplicates()
# print(len(unique_rows))

