import pandas as pd

source_file = '/Users/cmoore/Development/SampleData/Questions on Question Sets.xlsx'
target_file1 = '/Users/cmoore/Development/SampleData/industry_code_questions.xml'
target_file2 = '/Users/cmoore/Development/SampleData/industry_code_question_sets.xml'

questions = pd.read_excel(source_file)
#
# Create a copy of the Dataframe as the original will be used later to build the other xml files
# Remove rows with Duplicate ID and rows with duplicate Questions. Add New Column with the length of the question
#
questionCopy = questions.copy()
questionCopy.drop_duplicates(subset='Id', keep='first', inplace=True)
questionCopy.drop_duplicates(subset='Question', keep='first', inplace=True)
questionCopy['QuestionLength'] = questionCopy['Question'].map(len)
#
# Rename the Question column to QuestionText as this is what will be needed for the xml
#
questionCopy.rename(columns={'Question': 'QuestionText'}, inplace=True)
#
# Remove any questions that are too long and reset the index to the new Dataframe
#
xml = questionCopy[questionCopy['QuestionLength'] < 200].reset_index()
xml.drop_duplicates(subset='QuestionText', keep='first')
#
# Create the Public-id, Code and ValueType columns
#
xml['public-id'] = xml['Id']
xml['Code'] = xml['public-id']
question_check = xml['Id'].array
xml = xml.assign(ValueType='string')
#
# Generate the xml file and write the infomation to the target file
#
xml.to_xml(target_file1, index=False, root_name='import',
           row_name='IndustryCodeQuestion_Ext', attr_cols=['public-id'],
           elem_cols=['Code', 'QuestionText', 'ValueType'])

#
# Build arrays for each of the columns to be added to the DateFrame, where questions have been removed from the
# Industry Question these will be avoided in the Industry Question Set xml.
#
target_file = '/Users/cmoore/Development/SampleData/industry_code_question_sets.xml'
code = []
decline_question = []
referral_question = []
industry_code_question = []
effective_date = []
sequence = []
workflow_type = []
risk_type = []
process_value = []


def convert_workflow_type(workflow_type: str):
    if workflow_type.startswith('Turnover'):
        return 'Turnover'
    if workflow_type.startswith('Activity'):
        return 'Activity'
    if workflow_type.startswith('Qualifications'):
        return 'QualificationsExperience'


for question_row in questions.itertuples():
    pandas = question_row
    if not pandas.Id in question_check:
        continue

    risk_type_name = pandas.QuestionSetName.replace(' ', '')
    risk_type.append(risk_type_name)
    code.append(f'{risk_type_name}:{pandas.Id}')
    if pandas.Decline == 1:
        decline_question.append('true')
    else:
        decline_question.append('false')

    if pandas.Refer == 1:
        referral_question.append('true')
    else:
        referral_question.append('false')
    industry_code_question.append(pandas.Id)
    effective_date.append('2020-01-01 00:00:00.000')
    sequence.append(pandas.Level)
    workflow_type.append(convert_workflow_type(pandas.WorkflowName.replace(' ', '')))
    process_value.append(pandas.Rule)
#
# Create the Dataframe by adding each of the arrays. The arrays will form the columns in the Dataframe
#
question_set = pd.DataFrame({
    'public-id': code,
    'RiskType': risk_type,
    'Code': code,
    'DeclineQuestion': decline_question,
    'ReferralQuestion': referral_question,
    'EffectiveDate': effective_date,
    'Sequence': sequence,
    'WorkflowType': workflow_type,
    'IndustryCodeQuestion': industry_code_question,
    'ProcessValue': process_value
})
question_set.drop_duplicates(subset='public-id', keep='first', inplace=True)
#
# Convert the new Dataframe to xml, this does not write the xml to disk as we still need to make some changes
# to alow PolicyCenter to recognise the XML as the correct format for the System Table.
#
xml_string = question_set.to_xml(index=False,
                                 root_name='import',
                                 row_name='IndustryCodeQuestionSet_Ext', attr_cols=['public-id'],
                                 elem_cols=['Code', 'RiskType', 'DeclineQuestion', 'ReferralQuestion', 'EffectiveDate',
                                            'Sequence', 'WorkflowType', 'IndustryCodeQuestion', 'ProcessValue'])
#
# Rewrite the xml to have the correct format for Foreign Keys, a simpe replace is used
#
new_xml = xml_string
new_xml = new_xml.replace('<IndustryCodeQuestion>', '<IndustryCodeQuestion public-id="')
new_xml = new_xml.replace('</IndustryCodeQuestion>', '"/>')

with open(target_file2, 'w') as f:
    f.write(new_xml)
    f.flush()
