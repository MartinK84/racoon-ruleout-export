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
    age_cols = df.columns[list(df.loc['racoon-covid-19-demographic-information-age2',:].notna())]
    age_ecld = df.loc['racoon-covid-19-demographic-information-age2', age_cols].astype(int) >= 100
    df = df.drop(age_ecld.index[age_ecld], axis = 1)

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
    label_list = ['racoon-covid-19-abdominal-symptoms', 
                  'racoon-covid-19-annotation-pathology-atelectasis-scar-ancillary-feature-calciferous2', 
                  'racoon-covid-19-annotation-pathology-atelectasis-scar-ancillary-feature-curvilinear-pattern2', 
                  'racoon-covid-19-annotation-pathology-atelectasis-scar-ancillary-feature-radial-distribution2', 
                  'racoon-covid-19-annotation-pathology-atelectasis-scar-dominant-distribution2', 
                  'racoon-covid-19-annotation-pathology-atelectasis-scar-localization2', 
                  'racoon-covid-19-annotation-pathology-atelectasis-scar-severity', 
                  'racoon-covid-19-annotation-pathology-bronchi',
                  'racoon-covid-19-annotation-pathology-bronchiectasis-ancillary-feature-mucus-plugging2', 
                  'racoon-covid-19-annotation-pathology-bronchiectasis-ancillary-feature-with-traction2', 
                  'racoon-covid-19-annotation-pathology-bronchiectasis-dominant-distribution2', 
                  'racoon-covid-19-annotation-pathology-bronchiectasis-localization-lobes2', 
                  'racoon-covid-19-annotation-pathology-bronchiectasis-severity', 
                  'racoon-covid-19-annotation-pathology-bronchus-wall-thickening-dominant-distribution2', 
                  'racoon-covid-19-annotation-pathology-bronchus-wall-thickening-localization-lobes2', 
                  'racoon-covid-19-annotation-pathology-bronchus-wall-thickening-severity', 
                  'racoon-covid-19-annotation-pathology-cavitation-ancillary-feature-thin-walled-cystic', 
                  'racoon-covid-19-annotation-pathology-cavitation-ancillary-feature-with-air-crescent2', 
                  'racoon-covid-19-annotation-pathology-cavitation-ancillary-feature-with-air-fluid-level2', 
                  'racoon-covid-19-annotation-pathology-cavitation-ancillary-feature-with-halo2', 
                  'racoon-covid-19-annotation-pathology-cavitation-localization-lobes2', 
                  'racoon-covid-19-annotation-pathology-cavitation-severity', 
                  'racoon-covid-19-annotation-pathology-consolidation-ancillary-features-calciferous2', 
                  'racoon-covid-19-annotation-pathology-consolidation-ancillary-features-subpleural-sparing2', 
                  'racoon-covid-19-annotation-pathology-consolidation-ancillary-features-with-melting2', 
                  'racoon-covid-19-annotation-pathology-consolidation-dominant-distribution-anatomic2', 
                  'racoon-covid-19-annotation-pathology-consolidation-dominant-distribution-geographic2', 
                  'racoon-covid-19-annotation-pathology-consolidation-localization-lobes2', 
                  'racoon-covid-19-annotation-pathology-consolidation-severity', 
                  'racoon-covid-19-annotation-pathology-dominance-class', 
                  'racoon-covid-19-annotation-pathology-emphysema-ancillary-feature-bullous2', 
                  'racoon-covid-19-annotation-pathology-emphysema-ancillary-feature-paraseptal2', 
                  'racoon-covid-19-annotation-pathology-emphysema-dominant-pattern2', 
                  'racoon-covid-19-annotation-pathology-emphysema-localization-lobes2', 
                  'racoon-covid-19-annotation-pathology-emphysema-severity', 
                  'racoon-covid-19-annotation-pathology-ground-glass-region-ancillary-feature-subpleural-sparing2', 
                  'racoon-covid-19-annotation-pathology-ground-glass-region-ancillary-feature-with-consolidation-within-ground-glass2', 
                  'racoon-covid-19-annotation-pathology-ground-glass-region-ancillary-feature-with-crazy-paving2', 
                  'racoon-covid-19-annotation-pathology-ground-glass-region-ancillary-feature-with-reversed-halo2', 
                  'racoon-covid-19-annotation-pathology-ground-glass-region-ancillary-feature-with-vacuole-sign2', 
                  'racoon-covid-19-annotation-pathology-ground-glass-region-ancillary-feature-with-vessel-thickening-hyperemia2', 
                  'racoon-covid-19-annotation-pathology-ground-glass-region-dominant-distribution-anatomic2', 
                  'racoon-covid-19-annotation-pathology-ground-glass-region-dominant-distribution-geographic2', 
                  'racoon-covid-19-annotation-pathology-ground-glass-region-localization-lobes2', 
                  'racoon-covid-19-annotation-pathology-ground-glass-region-severity', 
                  'racoon-covid-19-annotation-pathology-lung-parenchyma', 
                  'racoon-covid-19-annotation-pathology-mass-larger-thirty-mm-ancillary-feature-calciferous2', 
                  'racoon-covid-19-annotation-pathology-mass-larger-thirty-mm-ancillary-feature-with-halo2', 
                  'racoon-covid-19-annotation-pathology-mass-larger-thirty-mm-ancillary-feature-with-infiltration-into-surroundings2', 
                  'racoon-covid-19-annotation-pathology-mass-larger-thirty-mm-ancillary-feature-with-melting2', 
                  'racoon-covid-19-annotation-pathology-mass-larger-thirty-mm-localization-lobes2', 
                  'racoon-covid-19-annotation-pathology-mass-larger-thirty-mm-severity', 
                  'racoon-covid-19-annotation-pathology-micronoduli-ancillary-feature-calciferous2', 
                  'racoon-covid-19-annotation-pathology-micronoduli-ancillary-feature-non-solid2', 
                  'racoon-covid-19-annotation-pathology-micronoduli-dominant-distribution-anatomic2', 
                  'racoon-covid-19-annotation-pathology-micronoduli-dominant-distribution-geographic2', 
                  'racoon-covid-19-annotation-pathology-micronoduli-localization-lobes2', 
                  'racoon-covid-19-annotation-pathology-micronoduli-severity2', 
                  'racoon-covid-19-annotation-pathology-nodule-ancillary-feature-calciferous2', 
                  'racoon-covid-19-annotation-pathology-nodule-ancillary-feature-irregular-with-halo', 
                  'racoon-covid-19-annotation-pathology-nodule-ancillary-feature-non-solid2', 
                  'racoon-covid-19-annotation-pathology-nodule-ancillary-feature-with-melting2', 
                  'racoon-covid-19-annotation-pathology-nodule-dominant-distribution-anatomic2', 
                  'racoon-covid-19-annotation-pathology-nodule-dominant-distribution-geographic2', 
                  'racoon-covid-19-annotation-pathology-nodule-localization-lobes2', 
                  'racoon-covid-19-annotation-pathology-nodule-severity2', 
                  'racoon-covid-19-annotation-pathology-pleura', 
                  'racoon-covid-19-annotation-pathology-pleural-disease-ancillary-feature-calciferous2', 
                  'racoon-covid-19-annotation-pathology-pleural-disease-dominant-distribution', 
                  'racoon-covid-19-annotation-pathology-pleural-disease-localization-laterality', 
                  'racoon-covid-19-annotation-pathology-pleural-effusion-hyperdense-greater-twenty-hu2', 
                  'racoon-covid-19-annotation-pathology-pleural-effusion-localization-laterality', 
                  'racoon-covid-19-annotation-pathology-pleural-effusion-severity', 
                  'racoon-covid-19-annotation-pathology-pleural-effusion-trapped2', 
                  'racoon-covid-19-annotation-pathology-reticulation-ancillary-feature-subpleural-sparing2', 
                  'racoon-covid-19-annotation-pathology-reticulation-ancillary-feature-with-honeycombing2', 
                  'racoon-covid-19-annotation-pathology-reticulation-dominant-distribution-anatomic2', 
                  'racoon-covid-19-annotation-pathology-reticulation-dominant-distribution-geographic2', 
                  'racoon-covid-19-annotation-pathology-reticulation-localization-lobes2', 
                  'racoon-covid-19-annotation-pathology-reticulation-severity', 
                  'racoon-covid-19-biochemical-alanine-aminotransferase', 
                  'racoon-covid-19-biochemical-alanine-aminotransferase-value', 
                  'racoon-covid-19-biochemical-albumin', 
                  'racoon-covid-19-biochemical-albumin-value', 
                  'racoon-covid-19-biochemical-alpha-hydroxybutyric-dehydrogenase', 
                  'racoon-covid-19-biochemical-alpha-hydroxybutyric-dehydrogenase-value', 
                  'racoon-covid-19-biochemical-aspartate-aminotransferase', 
                  'racoon-covid-19-biochemical-aspartate-aminotransferase-value', 
                  'racoon-covid-19-biochemical-cholinesterase', 
                  'racoon-covid-19-biochemical-cholinesterase-value', 
                  'racoon-covid-19-biochemical-creatine-kinase-muscle-brain-isoform', 
                  'racoon-covid-19-biochemical-creatine-kinase-muscle-brain-isoform-value', 
                  'racoon-covid-19-biochemical-creatinine', 
                  'racoon-covid-19-biochemical-creatinine-value2', 
                  'racoon-covid-19-biochemical-cystatin-c', 
                  'racoon-covid-19-biochemical-cystatin-c-value2', 
                  'racoon-covid-19-biochemical-egfr-value', 
                  'racoon-covid-19-biochemical-gamma-gt', 
                  'racoon-covid-19-biochemical-gamma-gt-value', 
                  'racoon-covid-19-biochemical-gfr', 
                  'racoon-covid-19-biochemical-globulin', 
                  'racoon-covid-19-biochemical-globulin-value', 
                  'racoon-covid-19-biochemical-glucose', 
                  'racoon-covid-19-biochemical-glucose-value2', 
                  'racoon-covid-19-biochemical-lactate', 
                  'racoon-covid-19-biochemical-lactate-dehydrogenase', 
                  'racoon-covid-19-biochemical-lactate-dehydrogenase-value', 
                  'racoon-covid-19-biochemical-lactate-value', 
                  'racoon-covid-19-biochemical-low-density-lipoprotein', 
                  'racoon-covid-19-biochemical-low-density-lipoprotein-value2',
                  'racoon-covid-19-biochemical-nt-pro-bnp', 
                  'racoon-covid-19-biochemical-nt-pro-bnp-value', 
                  'racoon-covid-19-biochemical-prealbumin', 
                  'racoon-covid-19-biochemical-prealbumin-value2', 
                  'racoon-covid-19-biochemical-total-bilirubin', 
                  'racoon-covid-19-biochemical-total-bilirubin-value', 
                  'racoon-covid-19-biochemical-total-protein', 
                  'racoon-covid-19-biochemical-total-protein-value', 
                  'racoon-covid-19-biochemical-troponin-i', 
                  'racoon-covid-19-biochemical-troponin-i-value', 
                  'racoon-covid-19-biochemical-troponin-t-value2', 
                  'racoon-covid-19-biochemical-troponin-t2', 
                  'racoon-covid-19-biochemical-urea', 
                  'racoon-covid-19-biochemical-urea-value2', 
                  'racoon-covid-19-bronchi-bronchiectasis-ancillary-feature-mucus-plugging2', 
                  'racoon-covid-19-bronchi-bronchiectasis-ancillary-feature-with-traction2', 
                  'racoon-covid-19-bronchi-bronchiectasis-dominant-distribution2', 
                  'racoon-covid-19-bronchi-bronchiectasis-localization-lobes2', 
                  'racoon-covid-19-bronchi-bronchiectasis-presence', 
                  'racoon-covid-19-bronchi-bronchiectasis-severity-left-lower-lobe', 
                  'racoon-covid-19-bronchi-bronchiectasis-severity-left-upper-lobe', 
                  'racoon-covid-19-bronchi-bronchiectasis-severity-lingula', 
                  'racoon-covid-19-bronchi-bronchiectasis-severity-right-lower-lobe', 
                  'racoon-covid-19-bronchi-bronchiectasis-severity-right-middle-lobe', 
                  'racoon-covid-19-bronchi-bronchiectasis-severity-right-upper-lobe', 
                  'racoon-covid-19-bronchi-bronchus-wall-thickening-dominant-distribution2', 
                  'racoon-covid-19-bronchi-bronchus-wall-thickening-localization-lobes2', 
                  'racoon-covid-19-bronchi-bronchus-wall-thickening-presence', 
                  'racoon-covid-19-bronchi-bronchus-wall-thickening-severity-left-lower-lobe', 
                  'racoon-covid-19-bronchi-bronchus-wall-thickening-severity-left-upper-lobe', 
                  'racoon-covid-19-bronchi-bronchus-wall-thickening-severity-lingula', 
                  'racoon-covid-19-bronchi-bronchus-wall-thickening-severity-right-lower-lobe', 
                  'racoon-covid-19-bronchi-bronchus-wall-thickening-severity-right-middle-lobe', 
                  'racoon-covid-19-bronchi-bronchus-wall-thickening-severity-right-upper-lobe', 
                  'racoon-covid-19-cardiac-disease', 
                  'racoon-covid-19-cardiac-disease-congestion', 
                  'racoon-covid-19-cardiac-symptoms', 
                  'racoon-covid-19-chronic-kidney-disease', 
                  'racoon-covid-19-chronic-kidney-disease-dialysis', 
                  'racoon-covid-19-coagulation-function-activated-partial-thromboplastin-time', 
                  'racoon-covid-19-coagulation-function-activated-partial-thromboplastin-time-value', 
                  'racoon-covid-19-coagulation-function-antithrombin-iii', 
                  'racoon-covid-19-coagulation-function-antithrombin-iii-value', 
                  'racoon-covid-19-coagulation-function-d-dimer', 
                  'racoon-covid-19-coagulation-function-d-dimer-value', 
                  'racoon-covid-19-coagulation-function-fibrinogen', 
                  'racoon-covid-19-coagulation-function-fibrinogen-value', 
                  'racoon-covid-19-coagulation-function-inr', 
                  'racoon-covid-19-coagulation-function-inr-value', 
                  'racoon-covid-19-cohort-primary-category', 
                  'racoon-covid-19-comorbidities-known', 
                  'racoon-covid-19-contact-to-infected-patients', 
                  'racoon-covid-19-copd', 
                  'racoon-covid-19-demographic-information-age2', 
                  'racoon-covid-19-demographic-information-gender', 
                  'racoon-covid-19-diabetes-insulin-therapy', 
                  'racoon-covid-19-diabetes-mellitus-presence-type', 
                  'racoon-covid-19-dominant-pathology-first', 
                  'racoon-covid-19-dominant-pathology-second', 
                  'racoon-covid-19-dominant-pathology-third', 
                  'racoon-covid-19-emphysem', 
                  'racoon-covid-19-emphysem-diagnosis-source', 
                  'racoon-covid-19-examination-ctdi-total2', 
                  'racoon-covid-19-examination-dlp-total2', 
                  'racoon-covid-19-examination-lung-image-acquisition-anatomical-regions', 
                  'racoon-covid-19-examination-lung-image-acquisition-anatomical-regions-including-abdomen-and-pelvis', 
                  'racoon-covid-19-examination-lung-image-acquisition-anatomical-regions-including-head-and-neckneck', 
                  'racoon-covid-19-examination-lung-image-acquisition-anatomical-regions-including-neck', 
                  'racoon-covid-19-examination-lung-image-acquisition-anatomical-regions-including-upper-abdomen', 
                  'racoon-covid-19-examination-lung-image-acquisition-contrast-phase-arterial', 
                  'racoon-covid-19-examination-lung-image-acquisition-contrast-phase-native',
                  'racoon-covid-19-examination-lung-image-acquisition-contrast-phase-portal-delayed-arteriovenous-mixed-phase', 
                  'racoon-covid-19-examination-lung-image-acquisition-contrast-phase-portal-delayed-venous', 
                  'racoon-covid-19-examination-lung-image-acquisition-contrast-phase-portal-venous', 
                  'racoon-covid-19-examination-lung-image-acquisition-contrast-phase-pulmonary-arterial-and-arterial-mixed', 
                  'racoon-covid-19-examination-lung-image-acquisition-contrast-phase-pulmonoarterial', 
                  'racoon-covid-19-examination-lung-image-acquisition-image-quality', 
                  'racoon-covid-19-examination-lung-image-acquisition-image-quality-beam-hardening-artifacts', 
                  'racoon-covid-19-examination-lung-image-acquisition-image-quality-incomplete-acquisition-of-target-region', 
                  'racoon-covid-19-examination-lung-image-acquisition-image-quality-limitation-due-to-image-noise',
                  'racoon-covid-19-examination-lung-image-acquisition-image-quality-motion-artifacts', 
                  'racoon-covid-19-examination-lung-image-acquisition2', 
                  'racoon-covid-19-examination-previous-modality', 
                  'racoon-covid-19-examination-previous-presence', 
                  'racoon-covid-19-fever2', 
                  'racoon-covid-19-hemoglobin', 
                  'racoon-covid-19-hemoglobin-value', 
                  'racoon-covid-19-hypertension', 
                  'racoon-covid-19-image-assessment-diagnosis', 
                  'racoon-covid-19-image-assessment-diagnosis-other', 
                  'racoon-covid-19-imaging-classification', 
                  'racoon-covid-19-imaging-classification-corads', 
                  'racoon-covid-19-imaging-classification-covrads', 
                  'racoon-covid-19-immunsuppresion', 
                  'racoon-covid-19-infection-related-indices-erythrocyte-sedimentation-rate', 
                  'racoon-covid-19-infection-related-indices-erythrocyte-sedimentation-rate-value', 
                  'racoon-covid-19-infection-related-indices-high-sensitivity-c-reactive-protein', 
                  'racoon-covid-19-infection-related-indices-high-sensitivity-c-reactive-protein-value', 
                  'racoon-covid-19-infection-related-indices-interleukin-six', 
                  'racoon-covid-19-infection-related-indices-interleukin-six-value', 
                  'racoon-covid-19-infection-related-indices-procalcitonin', 
                  'racoon-covid-19-infection-related-indices-procalcitonin-value', 
                  'racoon-covid-19-infection-related-indices-serum-ferritin', 
                  'racoon-covid-19-infection-related-indices-serum-ferritin-value2', 
                  'racoon-covid-19-liver-disease', 
                  'racoon-covid-19-lung-fibrosis', 
                  'racoon-covid-19-lung-fibrosis-diagnosis-source', 
                  'racoon-covid-19-lung-parenchyma-atelectasis-scar-ancillary-feature-calciferous2', 
                  'racoon-covid-19-lung-parenchyma-atelectasis-scar-ancillary-feature-curvilinear-pattern2', 
                  'racoon-covid-19-lung-parenchyma-atelectasis-scar-ancillary-feature-radial-distribution2', 
                  'racoon-covid-19-lung-parenchyma-atelectasis-scar-dominant-distribution2', 
                  'racoon-covid-19-lung-parenchyma-atelectasis-scar-localization2', 
                  'racoon-covid-19-lung-parenchyma-atelectasis-scar-presence', 
                  'racoon-covid-19-lung-parenchyma-atelectasis-scar-severity-left-lower-lobe', 
                  'racoon-covid-19-lung-parenchyma-atelectasis-scar-severity-left-upper-lobe', 
                  'racoon-covid-19-lung-parenchyma-atelectasis-scar-severity-lingula', 
                  'racoon-covid-19-lung-parenchyma-atelectasis-scar-severity-right-lower-lobe', 
                  'racoon-covid-19-lung-parenchyma-atelectasis-scar-severity-right-middle-lobe', 
                  'racoon-covid-19-lung-parenchyma-atelectasis-scar-severity-right-upper-lobe', 
                  'racoon-covid-19-lung-parenchyma-cavitation-ancillary-feature-thin-walled-cystic', 
                  'racoon-covid-19-lung-parenchyma-cavitation-ancillary-feature-with-air-crescent2', 
                  'racoon-covid-19-lung-parenchyma-cavitation-ancillary-feature-with-air-fluid-level2', 
                  'racoon-covid-19-lung-parenchyma-cavitation-ancillary-feature-with-halo2', 
                  'racoon-covid-19-lung-parenchyma-cavitation-localization-lobes2', 
                  'racoon-covid-19-lung-parenchyma-cavitation-presence', 
                  'racoon-covid-19-lung-parenchyma-cavitation-severity-left-lower-lobe', 
                  'racoon-covid-19-lung-parenchyma-cavitation-severity-left-upper-lobe', 
                  'racoon-covid-19-lung-parenchyma-cavitation-severity-lingula', 
                  'racoon-covid-19-lung-parenchyma-cavitation-severity-right-lower-lobe', 
                  'racoon-covid-19-lung-parenchyma-cavitation-severity-right-middle-lobe', 
                  'racoon-covid-19-lung-parenchyma-cavitation-severity-right-upper-lobe', 
                  'racoon-covid-19-lung-parenchyma-consolidation-ancillary-features-calciferous2', 
                  'racoon-covid-19-lung-parenchyma-consolidation-ancillary-features-subpleural-sparing2', 
                  'racoon-covid-19-lung-parenchyma-consolidation-ancillary-features-with-melting2', 
                  'racoon-covid-19-lung-parenchyma-consolidation-dominant-distribution-anatomic2', 
                  'racoon-covid-19-lung-parenchyma-consolidation-dominant-distribution-geographic2', 
                  'racoon-covid-19-lung-parenchyma-consolidation-localization-lobes2', 
                  'racoon-covid-19-lung-parenchyma-consolidation-presence', 
                  'racoon-covid-19-lung-parenchyma-consolidation-severity-left-lower-lobe', 
                  'racoon-covid-19-lung-parenchyma-consolidation-severity-left-upper-lobe', 
                  'racoon-covid-19-lung-parenchyma-consolidation-severity-lingula', 
                  'racoon-covid-19-lung-parenchyma-consolidation-severity-right-lower-lobe', 
                  'racoon-covid-19-lung-parenchyma-consolidation-severity-right-middle-lobe', 
                  'racoon-covid-19-lung-parenchyma-consolidation-severity-right-upper-lobe', 
                  'racoon-covid-19-lung-parenchyma-emphysema-ancillary-feature-bullous2', 
                  'racoon-covid-19-lung-parenchyma-emphysema-ancillary-feature-paraseptal2', 
                  'racoon-covid-19-lung-parenchyma-emphysema-dominant-pattern2', 
                  'racoon-covid-19-lung-parenchyma-emphysema-localization-lobes2', 
                  'racoon-covid-19-lung-parenchyma-emphysema-presence', 
                  'racoon-covid-19-lung-parenchyma-emphysema-severity-left-lower-lobe', 
                  'racoon-covid-19-lung-parenchyma-emphysema-severity-left-upper-lobe', 
                  'racoon-covid-19-lung-parenchyma-emphysema-severity-lingula', 
                  'racoon-covid-19-lung-parenchyma-emphysema-severity-right-lower-lobe', 
                  'racoon-covid-19-lung-parenchyma-emphysema-severity-right-middle-lobe', 
                  'racoon-covid-19-lung-parenchyma-emphysema-severity-right-upper-lobe', 
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-ancillary-feature-subpleural-sparing2', 
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-ancillary-feature-with-consolidation-within-ground-glass2', 
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-ancillary-feature-with-crazy-paving2', 
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-ancillary-feature-with-reversed-halo2', 
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-ancillary-feature-with-vacuole-sign2', 
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-ancillary-feature-with-vessel-thickening-hyperemia2', 
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-dominant-distribution-anatomic2', 
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-dominant-distribution-geographic2', 
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-localization-lobes2', 
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-presence', 
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-severity-left-lower-lobe', 
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-severity-left-upper-lobe', 
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-severity-lingula', 
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-severity-right-lower-lobe', 
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-severity-right-middle-lobe', 
                  'racoon-covid-19-lung-parenchyma-ground-glass-region-severity-right-upper-lobe', 
                  'racoon-covid-19-lung-parenchyma-mass-larger-thirty-mm-ancillary-feature-calciferous2', 
                  'racoon-covid-19-lung-parenchyma-mass-larger-thirty-mm-ancillary-feature-with-halo2', 
                  'racoon-covid-19-lung-parenchyma-mass-larger-thirty-mm-ancillary-feature-with-infiltration-into-surroundings2', 
                  'racoon-covid-19-lung-parenchyma-mass-larger-thirty-mm-ancillary-feature-with-melting2', 
                  'racoon-covid-19-lung-parenchyma-mass-larger-thirty-mm-localization-lobes2', 
                  'racoon-covid-19-lung-parenchyma-mass-larger-thirty-mm-presence', 
                  'racoon-covid-19-lung-parenchyma-mass-larger-thirty-mm-severity-left-lower-lobe', 
                  'racoon-covid-19-lung-parenchyma-mass-larger-thirty-mm-severity-left-upper-lobe', 
                  'racoon-covid-19-lung-parenchyma-mass-larger-thirty-mm-severity-lingula', 
                  'racoon-covid-19-lung-parenchyma-mass-larger-thirty-mm-severity-right-lower-lobe', 
                  'racoon-covid-19-lung-parenchyma-mass-larger-thirty-mm-severity-right-middle-lobe', 
                  'racoon-covid-19-lung-parenchyma-mass-larger-thirty-mm-severity-right-upper-lobe', 
                  'racoon-covid-19-lung-parenchyma-micronoduli-ancillary-feature-calciferous2', 
                  'racoon-covid-19-lung-parenchyma-micronoduli-ancillary-feature-non-solid2', 
                  'racoon-covid-19-lung-parenchyma-micronoduli-dominant-distribution-anatomic2', 
                  'racoon-covid-19-lung-parenchyma-micronoduli-dominant-distribution-geographic2', 
                  'racoon-covid-19-lung-parenchyma-micronoduli-localization-lobes2', 
                  'racoon-covid-19-lung-parenchyma-micronoduli-presence', 
                  'racoon-covid-19-lung-parenchyma-micronoduli-severity-left-lower-lobe', 
                  'racoon-covid-19-lung-parenchyma-micronoduli-severity-left-upper-lobe', 
                  'racoon-covid-19-lung-parenchyma-micronoduli-severity-lingula', 
                  'racoon-covid-19-lung-parenchyma-micronoduli-severity-right-lower-lobe', 
                  'racoon-covid-19-lung-parenchyma-micronoduli-severity-right-middle-lobe', 
                  'racoon-covid-19-lung-parenchyma-micronoduli-severity-right-upper-lobe', 
                  'racoon-covid-19-lung-parenchyma-nodule-ancillary-feature-calciferous2', 
                  'racoon-covid-19-lung-parenchyma-nodule-ancillary-feature-irregular-with-halo', 
                  'racoon-covid-19-lung-parenchyma-nodule-ancillary-feature-non-solid2', 
                  'racoon-covid-19-lung-parenchyma-nodule-ancillary-feature-with-melting2', 
                  'racoon-covid-19-lung-parenchyma-nodule-dominant-distribution-anatomic2', 
                  'racoon-covid-19-lung-parenchyma-nodule-dominant-distribution-geographic2', 
                  'racoon-covid-19-lung-parenchyma-nodule-localization-lobes2', 
                  'racoon-covid-19-lung-parenchyma-nodule-presence', 
                  'racoon-covid-19-lung-parenchyma-nodule-severity-left-lower-lobe', 
                  'racoon-covid-19-lung-parenchyma-nodule-severity-left-upper-lobe', 
                  'racoon-covid-19-lung-parenchyma-nodule-severity-lingula', 
                  'racoon-covid-19-lung-parenchyma-nodule-severity-right-lower-lobe', 
                  'racoon-covid-19-lung-parenchyma-nodule-severity-right-middle-lobe',
                  'racoon-covid-19-lung-parenchyma-nodule-severity-right-upper-lobe', 
                  'racoon-covid-19-lung-parenchyma-reticulation-ancillary-feature-subpleural-sparing2', 
                  'racoon-covid-19-lung-parenchyma-reticulation-ancillary-feature-with-honeycombing2', 
                  'racoon-covid-19-lung-parenchyma-reticulation-dominant-distribution-anatomic2', 
                  'racoon-covid-19-lung-parenchyma-reticulation-dominant-distribution-geographic2', 
                  'racoon-covid-19-lung-parenchyma-reticulation-localization-lobes2', 
                  'racoon-covid-19-lung-parenchyma-reticulation-presence', 
                  'racoon-covid-19-lung-parenchyma-reticulation-severity-left-lower-lobe', 
                  'racoon-covid-19-lung-parenchyma-reticulation-severity-left-upper-lobe', 
                  'racoon-covid-19-lung-parenchyma-reticulation-severity-lingula', 
                  'racoon-covid-19-lung-parenchyma-reticulation-severity-right-lower-lobe', 
                  'racoon-covid-19-lung-parenchyma-reticulation-severity-right-middle-lobe', 
                  'racoon-covid-19-lung-parenchyma-reticulation-severity-right-upper-lobe', 
                  'racoon-covid-19-lymphocytes', 
                  'racoon-covid-19-lymphocytes-value2', 
                  'racoon-covid-19-malignoma-metastatic-disease', 
                  'racoon-covid-19-mediastinum-aorta-sclerosis', 
                  'racoon-covid-19-mediastinum-atherosclerosis', 
                  'racoon-covid-19-mediastinum-lymphadenopathy-tumor-ancillary-feature-calciferous2', 
                  'racoon-covid-19-mediastinum-lymphadenopathy-tumor-localization', 
                  'racoon-covid-19-mediastinum-lymphadenopathy-tumor-presence', 
                  'racoon-covid-19-mediastinum-pericardial-effusion', 
                  'racoon-covid-19-monocytes', 
                  'racoon-covid-19-monocytes-value2', 
                  'racoon-covid-19-neurological-symptoms', 
                  'racoon-covid-19-neutrophils', 
                  'racoon-covid-19-neutrophils-value2', 
                  'racoon-covid-19-other-findings-information-details', 
                  'racoon-covid-19-other-findings-information-presence', 
                  'racoon-covid-19-oxygen-saturation', 
                  'racoon-covid-19-oxygenation-ratio', 
                  'racoon-covid-19-pack-years', 
                  'racoon-covid-19-patient-intubation-status-intubated', 
                  'racoon-covid-19-platelets', 
                  'racoon-covid-19-platelets-value2', 
                  'racoon-covid-19-pleura-pleural-disease-ancillary-feature-calciferous2', 
                  'racoon-covid-19-pleura-pleural-disease-dominant-distribution', 
                  'racoon-covid-19-pleura-pleural-disease-localization-laterality', 
                  'racoon-covid-19-pleura-pleural-disease-presence', 
                  'racoon-covid-19-pleura-pleural-effusion-hyperdense-greater-twenty-hu2', 
                  'racoon-covid-19-pleura-pleural-effusion-localization-laterality', 
                  'racoon-covid-19-pleura-pleural-effusion-presence', 
                  'racoon-covid-19-pleura-pleural-effusion-severity-left', 
                  'racoon-covid-19-pleura-pleural-effusion-severity-right', 
                  'racoon-covid-19-pleura-pleural-effusion-trapped2', 
                  'racoon-covid-19-pleura-pneumothorax-localization-laterality', 
                  'racoon-covid-19-pleura-pneumothorax-presence', 
                  'racoon-covid-19-pleura-pneumothorax-severity-left', 
                  'racoon-covid-19-pleura-pneumothorax-severity-right', 
                  'racoon-covid-19-pneumonia-disease-extent', 
                  'racoon-covid-19-pneumonia-disease-extent-severe', 
                  'racoon-covid-19-presenting-complaints-assessment-provide-information', 
                  'racoon-covid-19-respiratory-symptoms2', 
                  'racoon-covid-19-respiratoy-frequency', 
                  'racoon-covid-19-rt-pcr-assay3', 
                  'racoon-covid-19-sars-ct-score', 
                  'racoon-covid-19-sars-ct-score-lung-left-lower', 
                  'racoon-covid-19-sars-ct-score-lung-left-middle', 
                  'racoon-covid-19-sars-ct-score-lung-left-upper', 
                  'racoon-covid-19-sars-ct-score-lung-right-lower', 
                  'racoon-covid-19-sars-ct-score-lung-right-middle', 
                  'racoon-covid-19-sars-ct-score-lung-right-upper', 
                  'racoon-covid-19-sars-ct-score-total', 
                  'racoon-covid-19-smoking', 
                  'racoon-covid-19-specific-radiological-signs-bronchi-assessment', 
                  'racoon-covid-19-specific-radiological-signs-lung-parenchyma-assessment', 
                  'racoon-covid-19-specific-radiological-signs-pleura-assessment', 
                  'racoon-covid-19-supplementary-information-bronchi', 
                  'racoon-covid-19-supplementary-information-covid-specific', 
                  'racoon-covid-19-supplementary-information-lung-parenchyma', 
                  'racoon-covid-19-supplementary-information-mediastinum', 
                  'racoon-covid-19-supplementary-information-pleura', 
                  'racoon-covid-19-supplementary-information-vessels', 
                  'racoon-covid-19-systolic-pressure', 
                  'racoon-covid-19-vessels-arterial-occlusion', 
                  'racoon-covid-19-vessels-arterial-occlusion-calciferous2', 
                  'racoon-covid-19-vessels-arterial-occlusion-dominant-distribution', 
                  'racoon-covid-19-vessels-arterial-occlusion-localization-laterality', 
                  'racoon-covid-19-vessels-arterial-occlusion-severity', 
                  'racoon-covid-19-vessels-pulmonal-trunk-diameter-larger-than-aorta2', 
                  'racoon-covid-19-white-blood-cells', 
                  'racoon-covid-19-white-blood-cells-value2', 
                  'racoon-covid-19-xray-bronchi-bronchiectasis-dominant-distribution', 
                  'racoon-covid-19-xray-bronchi-bronchiectasis-localization', 
                  'racoon-covid-19-xray-bronchi-bronchiectasis-presence', 
                  'racoon-covid-19-xray-bronchi-bronchiectasis-severity-left', 
                  'racoon-covid-19-xray-bronchi-bronchiectasis-severity-right', 
                  'racoon-covid-19-xray-bronchi-bronchus-wall-thickening-dominant-distribution', 
                  'racoon-covid-19-xray-bronchi-bronchus-wall-thickening-localization', 
                  'racoon-covid-19-xray-bronchi-bronchus-wall-thickening-presence', 
                  'racoon-covid-19-xray-bronchi-bronchus-wall-thickening-severity-left', 
                  'racoon-covid-19-xray-bronchi-bronchus-wall-thickening-severity-right', 
                  'racoon-covid-19-xray-bronchi-pleural-disease-dominant-distribution', 
                  'racoon-covid-19-xray-bronchi-pleural-disease-localization', 
                  'racoon-covid-19-xray-bronchi-pleural-effusion-localization', 
                  'racoon-covid-19-xray-bronchi-pleural-effusion-severity-left', 
                  'racoon-covid-19-xray-bronchi-pleural-effusion-severity-right', 
                  'racoon-covid-19-xray-bronchi-pneumothorax-localization', 
                  'racoon-covid-19-xray-bronchi-pneumothorax-severity-left', 
                  'racoon-covid-19-xray-bronchi-pneumothorax-severity-right', 
                  'racoon-covid-19-xray-configuration-hilum-left', 
                  'racoon-covid-19-xray-configuration-hilum-right',
                  'racoon-covid-19-xray-diaphragmatic-elevation', 
                  'racoon-covid-19-xray-dominant-pathology-first', 
                  'racoon-covid-19-xray-dominant-pathology-second', 
                  'racoon-covid-19-xray-dominant-pathology-third', 
                  'racoon-covid-19-xray-external-material-cardiac', 
                  'racoon-covid-19-xray-external-material-central-venous-access2', 
                  'racoon-covid-19-xray-external-material-chest-tube-position', 
                  'racoon-covid-19-xray-external-material-ecmo', 
                  'racoon-covid-19-xray-external-material-endotracheal-tube-position', 
                  'racoon-covid-19-xray-external-material-nasogastric-tube-position2',
                  'racoon-covid-19-xray-external-material-others', 
                  'racoon-covid-19-xray-external-material-sternal-cerclage', 
                  'racoon-covid-19-xray-external-material-tracheal-cannula-position', 
                  'racoon-covid-19-xray-heart-mediastinum-cardiac-thoracic-quotient', 
                  'racoon-covid-19-xray-heart-mediastinum-soft-tissue-normal', 
                  'racoon-covid-19-xray-image-acquisition-information-examination-previous-modality', 
                  'racoon-covid-19-xray-image-acquisition-information-examination-previous-presence', 
                  'racoon-covid-19-xray-image-acquisition-information-image-quality', 
                  'racoon-covid-19-xray-image-acquisition-information-image-quality-acquisition-in-exspiration', 
                  'racoon-covid-19-xray-image-acquisition-information-image-quality-image-blurred', 
                  'racoon-covid-19-xray-image-acquisition-information-image-quality-incomplete-acquisition-of-target-region', 
                  'racoon-covid-19-xray-image-acquisition-information-image-quality-quality-image-rotated', 
                  'racoon-covid-19-xray-image-acquisition-information-image-quality-superposition-foreign-body', 
                  'racoon-covid-19-xray-image-acquisition-information-technique', 
                  'racoon-covid-19-xray-lung-parenchyma-atelectasis-scar-localization', 
                  'racoon-covid-19-xray-lung-parenchyma-atelectasis-scar-presence', 
                  'racoon-covid-19-xray-lung-parenchyma-atelectasis-scar-severity-dominant-distribution', 
                  'racoon-covid-19-xray-lung-parenchyma-atelectasis-scar-severity-left', 
                  'racoon-covid-19-xray-lung-parenchyma-atelectasis-scar-severity-right', 
                  'racoon-covid-19-xray-lung-parenchyma-cavitation-localization', 
                  'racoon-covid-19-xray-lung-parenchyma-cavitation-presence', 
                  'racoon-covid-19-xray-lung-parenchyma-cavitation-severity-left', 
                  'racoon-covid-19-xray-lung-parenchyma-cavitation-severity-right', 
                  'racoon-covid-19-xray-lung-parenchyma-consolidation-localization', 
                  'racoon-covid-19-xray-lung-parenchyma-consolidation-presence', 
                  'racoon-covid-19-xray-lung-parenchyma-consolidation-severity-dominant-distribution', 
                  'racoon-covid-19-xray-lung-parenchyma-consolidation-severity-left', 
                  'racoon-covid-19-xray-lung-parenchyma-consolidation-severity-right', 
                  'racoon-covid-19-xray-lung-parenchyma-emphysema-localization', 
                  'racoon-covid-19-xray-lung-parenchyma-emphysema-presence', 
                  'racoon-covid-19-xray-lung-parenchyma-emphysema-severity-dominant-distribution', 
                  'racoon-covid-19-xray-lung-parenchyma-emphysema-severity-left', 
                  'racoon-covid-19-xray-lung-parenchyma-emphysema-severity-right', 
                  'racoon-covid-19-xray-lung-parenchyma-ground-glass-region-localization', 
                  'racoon-covid-19-xray-lung-parenchyma-ground-glass-region-presence', 
                  'racoon-covid-19-xray-lung-parenchyma-ground-glass-region-severity-dominant-distribution', 
                  'racoon-covid-19-xray-lung-parenchyma-ground-glass-region-severity-left', 
                  'racoon-covid-19-xray-lung-parenchyma-ground-glass-region-severity-right', 
                  'racoon-covid-19-xray-lung-parenchyma-mass-larger-thirty-mm-localization', 
                  'racoon-covid-19-xray-lung-parenchyma-mass-larger-thirty-mm-presence',
                  'racoon-covid-19-xray-lung-parenchyma-mass-larger-thirty-mm-severity-left', 
                  'racoon-covid-19-xray-lung-parenchyma-mass-larger-thirty-mm-severity-right', 
                  'racoon-covid-19-xray-lung-parenchyma-micronodulus-localization', 
                  'racoon-covid-19-xray-lung-parenchyma-micronodulus-presence', 
                  'racoon-covid-19-xray-lung-parenchyma-micronodulus-severity-left', 
                  'racoon-covid-19-xray-lung-parenchyma-micronodulus-severity-right', 
                  'racoon-covid-19-xray-lung-parenchyma-nodule-localization2', 
                  'racoon-covid-19-xray-lung-parenchyma-nodule-presence2', 
                  'racoon-covid-19-xray-lung-parenchyma-nodule-severity-left', 
                  'racoon-covid-19-xray-lung-parenchyma-nodule-severity-right', 
                  'racoon-covid-19-xray-lung-parenchyma-reticulation-localization', 
                  'racoon-covid-19-xray-lung-parenchyma-reticulation-presence', 
                  'racoon-covid-19-xray-lung-parenchyma-reticulation-severity-dominant-distribution', 
                  'racoon-covid-19-xray-lung-parenchyma-reticulation-severity-left', 
                  'racoon-covid-19-xray-lung-parenchyma-reticulation-severity-right', 
                  'racoon-covid-19-xray-mediastinum-middle-positioned', 
                  'racoon-covid-19-xray-patient-intubation-status-intubated', 
                  'racoon-covid-19-xray-pleura-pleural-disease-presence', 
                  'racoon-covid-19-xray-pleura-pleural-effusion-presence', 
                  'racoon-covid-19-xray-pleura-pneumothorax-presence', 
                  'racoon-covid-19-xray-radiological-assessment-bronchi-assessment', 
                  'racoon-covid-19-xray-radiological-assessment-lung-parenchyma-assessment', 
                  'racoon-covid-19-xray-radiological-assessment-pleura-assessment', 
                  'racoon-covid-19-xray-soft-tissue-emphysema', 
                  'racoon-covid-19-xray-upper-mediastinum-enlarged', 
                  'racoon-covid-19-xray-vessels-apicobasal-distribution', 
                  'racoon-covid-19-xray-vessels-blurring', 
                  'racoon-covid-19-xray-vessels-caliber-change', 
                  'racoon-covid-19-treatment-protocol-antibiotics', 
                  'racoon-covid-19-treatment-protocol-antibiotics-azithromycin', 
                  'racoon-covid-19-treatment-protocol-antibiotics-carbapeneme', 
                  'racoon-covid-19-treatment-protocol-antibiotics-sublactam-ampicillin', 
                  'racoon-covid-19-treatment-protocol-antibiotics-sublactam-clarithromycin', 
                  'racoon-covid-19-treatment-protocol-antibiotics-tazobactam', 
                  'racoon-covid-19-treatment-protocol-chloroquine', 
                  'racoon-covid-19-treatment-protocol-chloroquine-details', 
                  'racoon-covid-19-treatment-protocol-hydroxychloroquine', 
                  'racoon-covid-19-treatment-protocol-lopinavir-ritonavir', 
                  'racoon-covid-19-treatment-protocol-methylprednisolone', 
                  'racoon-covid-19-treatment-protocol-oxygen-therapy', 
                  'racoon-covid-19-treatment-protocol-oxygen-therapy-type', 
                  'racoon-covid-19-treatment-protocol-remdesivir', 
                  'racoon-covid-19-treatment-protocol-thrombosis-prophylaxis', 
                  'racoon-covid-19-outcome-parameter-existing-signs-of-pulmonal-complications', 
                  'racoon-covid-19-outcome-parameter-last-documented-patient-outcome-description', 
                  'racoon-covid-19-outcome-parameter-worst-treatment-state-during-admission2']    
                   
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