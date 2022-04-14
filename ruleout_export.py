import xml.etree.ElementTree as et
from pathlib import Path
import pandas as pd
import hashlib
import argparse
import openpyxl
import uuid

# main program
def main(args):
    xml_file = Path(args.input)

    print(f"Loading input XML file {xml_file}")
    tree = et.parse(xml_file)
    root = tree.getroot()

    print(f"Getting list of all cases")
    cases = getAllCases(root)
    print(f"Found {len(cases)} cases")

    print(f"Parsing and anonymizing cases")
    case_assessment_list = []
    for case in cases:
        try:
            case_assessment = get_covid_assessment(case, args)
            case_assessment_list.append(case_assessment)
        except Exception as e:
            print(f"Error parsing case {case.attrib['CaseID']}")
            print(e)
            try:
                print("Error case info:")
                print(case.attrib)
                for c in case:
                    print(f"{c}: {c.attrib}")
            except:
                pass        

    print(f"Building output data")
    df = pd.DataFrame.from_records(case_assessment_list).transpose()

    # exclude data with age > 100
    try:
        age_cols = df.columns[list(df.loc['racoon-covid-19-demographic-information-age2',:].notna())]
        age_ecld = df.loc['racoon-covid-19-demographic-information-age2', age_cols].astype(int) >= 100
        df = df.drop(age_ecld.index[age_ecld], axis = 1)
    except:
        pass

    # save to file
    print(f"Writing output file: {args.output}")
    df.to_excel(args.output)

# get list of all available labels, including a unique list of the values
def get_label_list(cases, out_file, args):
    labels = []
    for case in cases:
        labels.append(get_covid_assessment_labels(case, args))
    labels_merge = {}
    for label in labels:
        for key in label.keys():
            if key in labels_merge.keys():
                labels_merge[key][1].append(label[key][1])
            else:
                labels_merge[key] = label[key]
                labels_merge[key][1] = [labels_merge[key][1]]
    for key in labels_merge.keys():
        labels_merge[key][1] = list(set(labels_merge[key][1]))

    df = pd.DataFrame.from_records(labels_merge).transpose()
    df.to_excel(out_file)

# parse information regarding covid assessment
def get_covid_assessment(case, args):
    label_list = ['racoon-covid-19-cohort-primary-category',
                  'racoon-covid-19-outcome-parameter-last-documented-patient-outcome-description',
                  'racoon-covid-19-outcome-parameter-worst-treatment-state-during-admission2',
                  'racoon-covid-19-outcome-parameter-existing-signs-of-pulmonal-complications',
                  'racoon-covid-19-treatment-protocol-oxygen-therapy',
                  'racoon-covid-19-treatment-protocol-lopinavir-ritonavir',
                  'racoon-covid-19-treatment-protocol-remdesivir',
                  'racoon-covid-19-treatment-protocol-antibiotics',
                  'racoon-covid-19-treatment-protocol-antibiotics-carbapeneme',
                  'racoon-covid-19-treatment-protocol-antibiotics-tazobactam',
                  'racoon-covid-19-treatment-protocol-antibiotics-sublactam-ampicillin',
                  'racoon-covid-19-treatment-protocol-antibiotics-sublactam-clarithromycin',
                  'racoon-covid-19-treatment-protocol-antibiotics-azithromycin',
                  'racoon-covid-19-treatment-protocol-oxygen-therapy-type',
                  'racoon-covid-19-treatment-protocol-thrombosis-prophylaxis',
                  'racoon-covid-19-treatment-protocol-methylprednisolone',
                  'racoon-covid-19-treatment-protocol-hydroxychloroquine',
                  'racoon-covid-19-treatment-protocol-chloroquine',
                  'racoon-covid-19-lung-parenchyma-emphysema-localization-lobes2',
                  'racoon-covid-19-lung-parenchyma-emphysema-ancillary-feature-paraseptal2',
                  'racoon-covid-19-lung-parenchyma-emphysema-dominant-pattern2',
                  'racoon-covid-19-lung-parenchyma-emphysema-ancillary-feature-bullous2',
                  'racoon-covid-19-lung-parenchyma-reticulation-localization-lobes2',
                  'racoon-covid-19-lung-parenchyma-reticulation-ancillary-feature-subpleural-sparing2',
                  'racoon-covid-19-lung-parenchyma-reticulation-ancillary-feature-with-honeycombing2',
                  'racoon-covid-19-lung-parenchyma-reticulation-dominant-distribution-geographic2',
                  'racoon-covid-19-lung-parenchyma-reticulation-dominant-distribution-anatomic2',
                  'racoon-covid-19-lung-parenchyma-cavitation-localization-lobes2',
                  'racoon-covid-19-lung-parenchyma-cavitation-ancillary-feature-with-halo2',
                  'racoon-covid-19-lung-parenchyma-cavitation-ancillary-feature-thin-walled-cystic',
                  'racoon-covid-19-lung-parenchyma-cavitation-ancillary-feature-with-air-fluid-level2',
                  'racoon-covid-19-lung-parenchyma-cavitation-ancillary-feature-with-air-crescent2',
                  'racoon-covid-19-lung-parenchyma-mass-larger-thirty-mm-localization-lobes2',
                  'racoon-covid-19-lung-parenchyma-mass-larger-thirty-mm-ancillary-feature-with-melting2',
                  'racoon-covid-19-lung-parenchyma-mass-larger-thirty-mm-ancillary-feature-with-halo2',
                  'racoon-covid-19-lung-parenchyma-mass-larger-thirty-mm-ancillary-feature-calciferous2',
                  'racoon-covid-19-lung-parenchyma-mass-larger-thirty-mm-ancillary-feature-with-infiltration-into-surroundings2',
                  'racoon-covid-19-lung-parenchyma-nodule-localization-lobes2',
                  'racoon-covid-19-lung-parenchyma-nodule-ancillary-feature-non-solid2',
                  'racoon-covid-19-lung-parenchyma-nodule-ancillary-feature-with-melting2',
                  'racoon-covid-19-lung-parenchyma-nodule-ancillary-feature-calciferous2',
                  'racoon-covid-19-lung-parenchyma-nodule-ancillary-feature-irregular-with-halo',
                  'racoon-covid-19-lung-parenchyma-nodule-dominant-distribution-geographic2',
                  'racoon-covid-19-lung-parenchyma-nodule-dominant-distribution-anatomic2',
                  'racoon-covid-19-lung-parenchyma-micronoduli-localization-lobes2',
                  'racoon-covid-19-lung-parenchyma-micronoduli-ancillary-feature-calciferous2',
                  'racoon-covid-19-lung-parenchyma-micronoduli-ancillary-feature-non-solid2',
                  'racoon-covid-19-lung-parenchyma-micronoduli-dominant-distribution-geographic2',
                  'racoon-covid-19-lung-parenchyma-micronoduli-dominant-distribution-anatomic2',
                  'racoon-covid-19-bronchi-bronchus-wall-thickening-localization-lobes2',
                  'racoon-covid-19-bronchi-bronchus-wall-thickening-dominant-distribution2',
                  'racoon-covid-19-bronchi-bronchiectasis-localization-lobes2',
                  'racoon-covid-19-bronchi-bronchiectasis-ancillary-feature-with-traction2',
                  'racoon-covid-19-bronchi-bronchiectasis-dominant-distribution2',
                  'racoon-covid-19-bronchi-bronchiectasis-ancillary-feature-mucus-plugging2',
                  'racoon-covid-19-pleura-pleural-effusion-hyperdense-greater-twenty-hu2',
                  'racoon-covid-19-pleura-pleural-effusion-trapped2',
                  'racoon-covid-19-patient-intubation-status-intubated',
                  'racoon-covid-19-pleura-pleural-disease-ancillary-feature-calciferous2',
                  'racoon-covid-19-vessels-arterial-occlusion',
                  'racoon-covid-19-vessels-arterial-occlusion-calciferous2',
                  'racoon-covid-19-vessels-pulmonal-trunk-diameter-larger-than-aorta2',
                  'racoon-covid-19-mediastinum-lymphadenopathy-tumor-presence',
                  'racoon-covid-19-mediastinum-lymphadenopathy-tumor-ancillary-feature-calciferous2',
                  'racoon-covid-19-mediastinum-pericardial-effusion',
                  'racoon-covid-19-mediastinum-atherosclerosis',
                  'racoon-covid-19-mediastinum-aorta-sclerosis',
                  'racoon-covid-19-sars-ct-score-total',
                  'racoon-covid-19-imaging-classification',
                  'racoon-covid-19-imaging-classification-corads',
                  'racoon-covid-19-imaging-classification-covrads',
                  'racoon-covid-19-annotation-pathology-lung-parenchyma',
                  'racoon-covid-19-lung-parenchyma-atelectasis-scar-localization2',
                  'racoon-covid-19-lung-parenchyma-atelectasis-scar-dominant-distribution2',
                  'racoon-covid-19-lung-parenchyma-atelectasis-scar-ancillary-feature-radial-distribution2',
                  'racoon-covid-19-lung-parenchyma-atelectasis-scar-ancillary-feature-curvilinear-pattern2',
                  'racoon-covid-19-lung-parenchyma-atelectasis-scar-ancillary-feature-calciferous2',
                  'racoon-covid-19-specific-radiological-signs-lung-parenchyma-assessment',
                  'racoon-covid-19-specific-radiological-signs-bronchi-assessment',
                  'racoon-covid-19-specific-radiological-signs-pleura-assessment',
                  'racoon-covid-19-lung-parenchyma-atelectasis-scar-presence',
                  'racoon-covid-19-lung-parenchyma-consolidation-presence',
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-presence',
                  'racoon-covid-19-lung-parenchyma-emphysema-presence',
                  'racoon-covid-19-lung-parenchyma-reticulation-presence',
                  'racoon-covid-19-lung-parenchyma-cavitation-presence',
                  'racoon-covid-19-lung-parenchyma-mass-larger-thirty-mm-presence',
                  'racoon-covid-19-lung-parenchyma-nodule-presence',
                  'racoon-covid-19-lung-parenchyma-micronoduli-presence',
                  'racoon-covid-19-bronchi-bronchus-wall-thickening-presence',
                  'racoon-covid-19-bronchi-bronchiectasis-presence',
                  'racoon-covid-19-pleura-pneumothorax-presence',
                  'racoon-covid-19-pleura-pleural-effusion-presence',
                  'racoon-covid-19-pleura-pleural-disease-presence',
                  'racoon-covid-19-lung-parenchyma-consolidation-localization-lobes2',
                  'racoon-covid-19-lung-parenchyma-consolidation-ancillary-features-subpleural-sparing2',
                  'racoon-covid-19-lung-parenchyma-consolidation-dominant-distribution-geographic2',
                  'racoon-covid-19-lung-parenchyma-consolidation-dominant-distribution-anatomic2',
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-localization-lobes2',
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-ancillary-feature-with-consolidation-within-ground-glass2',
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-ancillary-feature-with-vessel-thickening-hyperemia2',
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-ancillary-feature-subpleural-sparing2',
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-ancillary-feature-with-crazy-paving2',
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-ancillary-feature-with-reversed-halo2',
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-ancillary-feature-with-vacuole-sign2',
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-dominant-distribution-geographic2',
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-dominant-distribution-anatomic2',
                  'racoon-covid-19-demographic-information-age2',
                  'racoon-covid-19-demographic-information-gender',
                  'racoon-covid-19-contact-to-infected-patients',
                  'racoon-covid-19-emphysem',
                  'racoon-covid-19-copd',
                  'racoon-covid-19-hypertension',
                  'racoon-covid-19-cardiac-disease',
                  'racoon-covid-19-cardiac-disease-congestion',
                  'racoon-covid-19-liver-disease',
                  'racoon-covid-19-chronic-kidney-disease',
                  'racoon-covid-19-chronic-kidney-disease-dialysis',
                  'racoon-covid-19-diabetes-mellitus-presence-type',
                  'racoon-covid-19-diabetes-insulin-therapy',
                  'racoon-covid-19-lung-fibrosis',
                  'racoon-covid-19-comorbidities-known',
                  'racoon-covid-19-immunsuppresion',
                  'racoon-covid-19-malignoma-metastatic-disease',
                  'racoon-covid-19-smoking',
                  'racoon-covid-19-pack-years',
                  'racoon-covid-19-abdominal-symptoms',
                  'racoon-covid-19-cardiac-symptoms',
                  'racoon-covid-19-fever2',
                  'racoon-covid-19-respiratoy-frequency',
                  'racoon-covid-19-systolic-pressure',
                  'racoon-covid-19-oxygen-saturation',
                  'racoon-covid-19-respiratory-symptoms2',
                  'racoon-covid-19-neurological-symptoms',
                  'racoon-covid-19-rt-pcr-assay3',
                  'racoon-covid-19-monocytes',
                  'racoon-covid-19-platelets',
                  'racoon-covid-19-hemoglobin',
                  'racoon-covid-19-white-blood-cells',
                  'racoon-covid-19-neutrophils',
                  'racoon-covid-19-lymphocytes',
                  'racoon-covid-19-biochemical-total-protein',
                  'racoon-covid-19-biochemical-albumin',
                  'racoon-covid-19-biochemical-globulin',
                  'racoon-covid-19-biochemical-prealbumin',
                  'racoon-covid-19-biochemical-urea',
                  'racoon-covid-19-biochemical-total-bilirubin',
                  'racoon-covid-19-biochemical-creatinine',
                  'racoon-covid-19-biochemical-gfr',
                  'racoon-covid-19-biochemical-glucose',
                  'racoon-covid-19-biochemical-creatine-kinase-muscle-brain-isoform',
                  'racoon-covid-19-biochemical-cholinesterase',
                  'racoon-covid-19-biochemical-cystatin-c',
                  'racoon-covid-19-biochemical-lactate',
                  'racoon-covid-19-biochemical-lactate-dehydrogenase',
                  'racoon-covid-19-biochemical-alpha-hydroxybutyric-dehydrogenase',
                  'racoon-covid-19-biochemical-low-density-lipoprotein',
                  'racoon-covid-19-biochemical-gamma-gt',
                  'racoon-covid-19-biochemical-troponin-t2',
                  'racoon-covid-19-biochemical-troponin-i',
                  'racoon-covid-19-biochemical-nt-pro-bnp',
                  'racoon-covid-19-biochemical-aspartate-aminotransferase',
                  'racoon-covid-19-biochemical-alanine-aminotransferase',
                  'racoon-covid-19-infection-related-indices-serum-ferritin',
                  'racoon-covid-19-infection-related-indices-high-sensitivity-c-reactive-protein',
                  'racoon-covid-19-infection-related-indices-interleukin-six',
                  'racoon-covid-19-infection-related-indices-procalcitonin',
                  'racoon-covid-19-infection-related-indices-erythrocyte-sedimentation-rate',
                  'racoon-covid-19-coagulation-function-d-dimer',
                  'racoon-covid-19-coagulation-function-activated-partial-thromboplastin-time',
                  'racoon-covid-19-coagulation-function-fibrinogen',
                  'racoon-covid-19-coagulation-function-antithrombin-iii',
                  'racoon-covid-19-coagulation-function-inr']
                   
    covid_assessment = {}

    # fails for unknown reasons for some sites
    lastname = ''
    try:
        lastname = case[0].attrib['LastName']
    except:
        print(f"Error getting lastname for case {case.attrib['CaseID']}, trying to continue without")
        pass
    
    # if building the case_string fails use a random uuid
    try:
        case_string = case.attrib['CaseID'] + lastname + case[0].attrib['PatientID'] + case[0].attrib['InstitutionName']
        hash_string = encrypt(case_string)
    except:
        print(f"Error building case_string for case {case.attrib['CaseID']}, using random UUID")
        hash_string = str(uuid.uuid4())
    covid_assessment["ID"] = hash_string

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
                                    if args.verbose:
                                        print(f"{question_type} already in list with value: {old_value} vs. {new_value}, using {old_value}")
                                    new_value = old_value                             
                        covid_assessment[question_type] = new_value
                else:
                    print(f"No type attribute in question {question.attrib[pair[0]]}")
        
    return covid_assessment

# parse table of all possible assessments, including internal type, human readable text and example string
def get_covid_assessment_labels(case, args):                    
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
                                if args.verbose:
                                    print(f"{question_type} already in list with value: {old_value} vs. {new_value}, using {old_value}")
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
        import tkinter
        from tkinter import filedialog as fd

        filetypes = (
                ('XML files', '*.xml'),
                ('All files', '*.*')
            )

        root = tkinter.Tk()
        root.wm_withdraw() # this completely hides the root window

        filename = fd.askopenfilename(
            title='Select input file',
            initialdir=Path(__file__).parent,
            filetypes=filetypes
            )
        root.destroy()

        args.input = Path(filename)

    if args.output is None:
        args.output = Path.joinpath(Path(args.input).parent, Path(args.input).stem + ".xlsx")

    main(args)