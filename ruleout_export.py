import xml.etree.ElementTree as et
from pathlib import Path
import pandas as pd
import hashlib
import argparse

# main program
def main(args):
    xml_file = Path(args.input)

    print(f"Loading input XML file {xml_file}")
    tree = et.parse(xml_file)
    root = tree.getroot()

    print(f"Getting list of all cases")
    cases = getAllCases(root)
    print(f"\tFound {len(cases)} cases")

    print(f"Parsing cases")
    case_assessment_list = []
    for case in cases:
        case_assessment = get_covid_assessment(case)
        case_assessment_list.append(case_assessment)

    print(f"List of assessment labels:")
    print(get_covid_assessment_labels(cases[0]))

# parse information regarding covid assessment
def get_covid_assessment(case):
    covid_assessment = {}
    label_list = ["racoon-covid-19-annotation-pathology-lung-parenchyma"] # ToDo: complete list of labels for ruleout
                   
    covid_assessment = {}
    question_pairs = [('Label', 'QuestionType'), ('Question', 'Type')]
    for question in case.iter('Question'):
        for pair in question_pairs:
            if pair[0] in question.attrib:
                if pair[1] in question.attrib:
                    question_type = question.attrib[pair[1]]
                    new_value = question.attrib['Answer']

                    if question_type in label_list:
                        # check if question type already exists in list, possibly from other question/label 
                        if question_type in covid_assessment.keys():
                            # check if existing value and new value are not the same
                            old_value = covid_assessment[question_type]
                            if old_value != new_value:                            
                                if len(new_value) == 0: # keep old value if new value is invalid
                                    print(f"{question_type} already in list with value: {old_value} vs. {new_value}")
                                    new_value = old_value                             
                        covid_assessment[question_type] = new_value
                else:
                    print(f"No type attribute in question {question.attrib[pair[0]]}")
        
    return covid_assessment

# parse table of all possible assessments, including internal type, human readable text and example string
def get_covid_assessment_labels(case):                    
    covid_assessment = {}
    question_pairs = [('Label', 'QuestionType'), ('Question', 'Type')]
    for question in case.iter('Question'):
        for pair in question_pairs:
            if pair[0] in question.attrib:
                if pair[1] in question.attrib:
                    question_type = question.attrib[pair[1]]
                    new_value = question.attrib['Answer']

                    # check if question type already exists in list, possibly from other question/label 
                    if question_type in covid_assessment.keys():
                        # check if existing value and new value are not the same
                        old_value = covid_assessment[question_type][1]
                        if old_value != new_value:                            
                            if len(new_value) == 0: # keep old value if new value is invalid
                                print(f"{question_type} already in list with value: {old_value} vs. {new_value}")
                                new_value = old_value                                                             
                    covid_assessment[question_type] = [question.attrib[pair[0]], new_value]
                else:
                    print(f"No type attribute in question {question.attrib[pair[0]]}")
        
    return covid_assessment

# use sha256 hash for anonymization
def encrypt(string):
    hash_object = hashlib.sha256(string.encode())
    hash_string = hash_object.hexdigest()
    return hash_string

# parse cases as list
def getAllCases(root) -> list:
    cases = []
    for trial in root.iter('Trial'):
        for trialArm in trial.iter('TrialArm'):
            for case in trialArm.iter('Case'):
                cases.append(case)
    return cases

# program entry point
if __name__ == "__main__":
    # define arguments
    parser = argparse.ArgumentParser(description="Anonymize and export parameters for RACOON Ruleout from Mint XML dump")
    parser.add_argument("-i", "--input", help="Input XML file")
    parser.add_argument("-o", "--output", help="Output Excel file")    
    parser.add_argument("-v", "--verbose", help="Verbose log output", action="store_true")

    args = parser.parse_args()

    if args.input is None:
        args.input = Path("MintExportRACOON.xml")
    if args.output is None:
        args.output = Path("MintExportRACOON.xlsx")

    main(args)