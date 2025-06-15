# 2024_Early_Warning_SakuMetall
Early warning system for improving supply reliability at Saku Metall
*This is a template repository for this organization. Start by replacing the placeholder for the project name with its actual title.*

## Summary
| Company Name | [SakuMetall](https://sakumetall.ee/) |
| :--- | :--- |
| Development Team Lead Name | Tara Ghasempouri |
| Development Team Lead E-mail | tara.ghasempouri@taltech.ee |
| Duration of the Demonstration Project | 01/09/2024–30/05/2025 |
| Final Report | [Example_report.pdf](https://github.com/ai-robotics-estonia/_project_template_/files/13800685/IC-One-Page-Project-Status-Report-10673_PDF.pdf) |

### Each project has an alternative for documentation
1. Fill in the [description](#description) directly in the README below *OR*;
2. make a [custom agreement with the AIRE team](#custom-agreement-with-the-AIRE-team).

# Description
## Objectives of the Demonstration Project
The aim of the project was to improve the order receiving and processing process for Saku Metall Allhanke Tehas AS’s (SMAHT) largest client, KONE. The impact of order receiving and processing process carries over to the entire production chain and affects the efficiency of resource utilization, quality, and most importantly, supply reliability. 
The broadest goal of the demo project was to increase supply reliability, but there were many smaller goals as well, that are similarly important. For example: reducing the possibility of human error, reducing repetitive (data management) work, increasing employee productivity, optimizing the purchasing process and inventory, increasing material utilization efficiency, reducing waste generation, and improving customer satisfaction.
The result of the demo project was to develop an artificial intelligence solution that enables:
    • Obtain information about unique behavior in multi-level Bill of Material (BOM). Unique behavior occurs when BOM consists of unique parts, material and purchased components which are used rarely.
    
## Activities and Results of the Demonstration Project
### Challenge
Saku Metall Project Summary:
Saku Metall has accumulated a large Bill of Materials (BOM) over many years of purchases. The primary goal of this project was to identify unique patterns—specifically, rare purchases—within this extensive dataset.
Challenges:
After several discussions with the company and a thorough analysis of the dataset, we identified two main variables that are critical for pattern detection. However, the number of unique values for these two variables could reach into the millions. Initially, it was assumed that 7–8 variables would be analyzed, each with only thousands of values. Therefore, the AI tool had to be customized to handle a much larger and more complex dataset than originally anticipated.
During the course of the project, we also discovered that the dataset is updated every three months. In fact, new data is continuously added to the historical dataset, which can significantly alter the initial analysis results. The original assumption was that the dataset would be static. As a result, the AI workflow had to be adapted to accommodate these regular updates and ensure consistent results over time.
Additionally, we found that the AI algorithms we initially used—namely the Apriori algorithm and FP-Growth, both well-known for association rule mining—were not well-suited to the specific nature of this dataset. These algorithms rely on a tokenization process that proved to be extremely resource-intensive due to redundant operations that were not necessary for our use case.
Based on the Development team’s in-depth understanding, it became clear that optimizing the tokenization step was essential. We reviewed the inner workings of both algorithms and implemented customized modifications to significantly improve their efficiency. These adjustments made the process much less resource-hungry while maintaining high accuracy in identifying rare and unique patterns in the BOM.
This approach differed substantially from the initial plan, which assumed the AI algorithms could be applied directly without customization.

### Data Sources
The tool operates using three primary input files, all placed under the /Input/ folder:

1. Historical_BOM.zip: A collection of archived Excel BOM files gathered from previous orders and projects. These represent the historical baseline used to compute frequency-based rarity metrics.
2. To_be_Added.zip: A ZIP archive containing new BOM Excel files to be processed and compared against the historical baseline. After each run:

New component–material pairs are compared against Sheet 5 (metrics)
Novel entries are flagged as New
This data is then merged into the historical archive

3. KONE_Critical_Item.xlsx: This Excel file includes a list of critical components, manually maintained by domain experts or derived from risk analysis.
The items mentioned in this folder receive a Critical_Flag = Critical during analysis, ensuring they are highlighted even if they occur frequently. Non-critical components are flagged as Safe.


### AI Technologies
This project uses Association Rule Mining (ARM) concepts, specifically:

Support: How often a (Component, Material) pair occurs in all files
Confidence: How likely a Material is seen given a Component
Support + Confidence: Combined metric used for enhanced prioritization

Techniques Applied:
Hybrid Frequency Modeling: Combines Apriori-style pair counting with FP-Growth-inspired optimizations

Deduplication: Ensures frequency is counted based on unique BOM files, not line repetitions

Rule-based classification:
New: if not seen in historical data
Rare: if support < 3.5%
Critical: if component is in critical list

### Technological Results
Key Functionalities:
Reads historical and new BOMs from .zip archives
Parses and normalizes component/material IDs

Flags:
Rare pairs (low support)
New pairs (not in historical)
Critical components (from separate file)

Updates and saves 5 Excel sheets:
1_Historical: Deduplicated historical data with metrics and criticality status
2_To_Be_Added: Newly added data, flagged as New, Rare, or Critical
3_Merged: Unified view of all (highlighting new entries in green)
4_Metrics: ARM results: Support, Confidence, and their sum
5_Total_Count: Final clean count of all Component–Material combinations, merged and updated

Automation Benefits: Merges To_be_Added.zip content into historical automatically

After execution, new data is fully absorbed into baseline Prevents re-analysis or duplication Ensures all metrics are updated for the next cycle

### Technical Architecture
The system follows a modular pipeline for processing and analyzing BOM data. Below is a breakdown of each component and their interactions:
1. Input Handling Module
Component: ZIP Extractor & Normalizer
Function: Reads .zip files (Historical_BOM.zip, To_be_Added.zip) and extracts valid .xlsx BOM sheets.
Key Task: Cleans and normalizes Component and Material fields using regex (KM\d+ pattern).
Integration: Takes raw engineering BOMs exported from Saku Metall's internal systems or ERP software.

2. Critical Item Flagger
Component: Critical Reference Checker
Function: Compares all Component values against KONE_Critical_Item.xlsx to tag critical components.
Output: Adds a Critical_Flag column (Critical / Safe) to both historical and new entries.

3. Association Rule Metrics Calculator
Component: ARM Metrics Engine
Technology: ARM logic inspired by Apriori and FP-Growth.

Calculations:
Support: Pair frequency across all BOM files
Confidence: Material likelihood given a Component
Support_Confidence_Sum: Composite prioritization metric

4. Classification Engine
Component: Rule-Based Classifier
Logic:
New: Pair not found in historical archive
Rare: Support < 3.5%
Critical: Found in critical reference file
Usage: Automatically annotates BOM entries for procurement or QA review.

5. Historical Archiver and Merging Engine
Component: Data Merger & Deduplicator
Function: After processing To_be_Added.zip, the new entries are:
Compared with existing records
Merged into historical archive
Deleted from incoming ZIPs to avoid reprocessing
Integration: Ensures the dataset evolves over time and supports continuous improvement.

6. Excel Reporting Generator
Component: Multi-Sheet Excel Writer
Libraries: pandas, openpyxl

Sheets Generated:
1_Historical: Historical deduplicated records
2_To_Be_Added: New entries with status flags
3_Merged: All current BOM data with highlights
4_Metrics: ARM output for each pair
5_Total_Count: Clean count summary

(Visual Architecture Diagram)
                       +------------------------------+
                       |        Input Data Files       |
                       |------------------------------|
                       | 1. Historical_BOM.zip         |
                       | 2. To_be_Added.zip            |
                       | 3. KONE_Critical_Item.xlsx    |
                       +------------------------------+
                                    |
                                    v
                     +-------------------------------+
                     |  BOM Normalization & Parsing  |
                     +-------------------------------+
                              |           |
                              v           v
               +------------------+     +----------------------+
               | Historical BOMs  |     |   To_be_Added BOMs   |
               +------------------+     +----------------------+
                              \           /
                               \         /
                                \       /
                                 v     v
                     +-------------------------------+
                     |   Unified Component-Material   |
                     |     Formatting & Deduplication |
                     +-------------------------------+
                                    |
                                    v
                     +-------------------------------+
                     |    Critical Flag Assignment    |
                     | (using KONE_Critical_Item.xlsx)|
                     +-------------------------------+
                                    |
                                    v
                     +-------------------------------+
                     |     ARM Metrics Calculation    |
                     |     (Support, Confidence)      |
                     +-------------------------------+
                                    |
                                    v
                     +-------------------------------+
                     |   Classification & Labeling    |
                     |   (Rare, New, Not Rare, etc.)  |
                     +-------------------------------+
                                    |
                                    v
                     +-------------------------------+
                     | Excel Export with Highlights   |
                     |  (Color-coded for usability)   |
                     +-------------------------------+
                                    |
                                    v
           +--------------------------------------------------+
           | Updated Artifacts: Merged ZIP + Excel Reports     |
           | (Historical, To_be_Added, Metrics, Summary, Rules)|
           +--------------------------------------------------+


### User Interface 
The solution is implemented as a command-line script. Users interact with the system by running Python scripts that:
- Process Excel BOM files from `.zip` archives
- Automatically flag rare, critical, or new entries
- Export annotated results in Excel format

The output is designed to be human-readable, color-coded, and can be integrated into Excel-based workflows or uploaded to internal ERP systems for further review.

### Future Potential of the Technical Solution
- Real-time decision support: With minimal enhancements, the AI engine can operate in near real-time, helping planners prioritize procurement dynamically.
- Predictive procurement planning: Combining ARM with machine learning forecasting techniques can help prevent disruptions before they occur.
- Cloud-native scalability: The scripts can be migrated to a serverless architecture or cloud pipelines (e.g., AWS Lambda, Azure Functions) for scalable, on-demand analysis.
- User-facing dashboards: Integrating output into BI tools (e.g., Power BI or Tableau) would make insights more accessible and interactive for supply chain teams.
- Risk scoring system: Patterns detected could be translated into a numeric risk score for each order or supplier, allowing for smarter resource allocation.

### Lessons Learned
- Static assumptions about datasets rarely hold in real industry scenarios.
- Highly customized AI solutions outperform generic ARM libraries.
- Close feedback loops with domain experts improve both performance and usability.
- Rule-based interpretability is essential in manufacturing.

# Custom agreement with the AIRE team
*If you have a unique project or specific requirements that don't fit neatly into the Docker file or description template options, we welcome custom agreements with our AIRE team. This option allows flexibility in collaborating with us to ensure your project's needs are met effectively.*

*To explore this option, please contact our demonstration projects service manager via katre.eljas@taltech.ee with the subject line "Demonstration Project Custom Agreement Request - [Your Project Name]." In your email, briefly describe your project and your specific documentation or collaboration needs. Our team will promptly respond to initiate a conversation about tailoring a solution that aligns with your project goals.*
