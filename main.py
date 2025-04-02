from core import PPTUpdater

def main():
    template_path = "template.pptx"
    data_path = "inspirient_analysis_data_prorotype_v02-GA.json"
    output_path = "updated_presentation.pptx"
    
    try:
        updater = PPTUpdater(template_path, data_path)
        if updater.table_data:
            updater.update_slides()
            updater.save(output_path)
        else:
            updater.logger.error("No table data available. Skipping update.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()
####start123