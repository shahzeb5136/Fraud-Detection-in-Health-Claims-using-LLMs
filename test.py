import pandas as pd
import ollama
import openai
import re
import os
import sys
from openpyxl import load_workbook

# -----------------------------
# 0) Constants & Configuration
# -----------------------------

# --- Agent Prompts Definition ---
# Define each agent's system prompt here.
# The key will be used to name the output column (spaces replaced with '_').
AGENT_PROMPTS = {
    "Agent_1_Diagnosis_Service_Match": """
You are an AI agent focused on the consistency between Diagnosis and Services Provided in a medical claim.
Based *only* on the information given for this claim (Claim ID, Diagnosis, Services Provided), assess how well the Services align with what might typically be expected for the stated Diagnosis.
Assign a score from 1 (Highly Consistent - Low Concern) to 10 (Highly Inconsistent - High Concern).
Provide a single sentence explaining your reasoning for the score, focusing *only* on the diagnosis-service relationship.
Output Format:
Score: [1-10]
Rationale: [Your one-sentence explanation]
""",
    "Agent_2_Upcoding_Check": """
You are an AI agent specializing in identifying potential upcoding in medical claims.
Analyze the provided Services. Do they seem excessively complex or comprehensive given the Diagnosis?
Consider if simpler, less costly services might have been appropriate.
Provide a brief text assessment (1-2 sentences) indicating whether there's a low, medium, or high suspicion of upcoding based *only* on the Diagnosis and Services list. Explain your reasoning concisely.
Output Format:
Upcoding Suspicion: [Low/Medium/High]
Reasoning: [Your brief explanation]
""",
    "Agent_3_Service_Frequency_Alert": """
You are an AI agent evaluating the frequency or duration of services provided for a given diagnosis.
Based on the Diagnosis and Services Provided, does the *quantity* or *duration* (if implied) of services seem unusual or excessive for treating the stated diagnosis in a typical scenario?
Focus solely on the volume/frequency aspect.
State whether the frequency appears 'Normal' or 'Potentially Excessive'. Add a one-sentence justification.
Output Format:
Frequency Assessment: [Normal/Potentially Excessive]
Justification: [Your one-sentence justification]
""",
    # --- ADD MORE AGENTS HERE ---
    # Give each agent a unique, descriptive name (key) and its specific system prompt (value).
    # "Agent_4_Your_Specific_Check_Name": """
    # Your specific instructions for this agent...
    # Define the expected output format if desired.
    # """,
    # ... Add up to Agent_100 or more as needed ...
    # "Agent_100_Final_Review_Check": """
    # Instructions for the 100th agent...
    # """
}
# -----------------------------

OLLAMA_MODEL = "llama3.1"  # Or whichever model you have
OPENAI_MODEL = "gpt-3.5-turbo" # Or "gpt-4", "gpt-4o-mini", etc.
INPUT_EXCEL_FILE = "Multiple Claims.xlsx"
OUTPUT_EXCEL_FILE = "Multiple Claims - Agentic Output.xlsx" # Changed output filename
EXCEL_SHEET_NAME = "Sheet1"

# -----------------------------
# 1) SETUP & USER CHOICE (Unchanged from your original)
# -----------------------------
def get_api_choice():
    """Prompts the user to select the API to use."""
    while True:
        print("Select the AI model/API to use:")
        print("1: Ollama (using {})".format(OLLAMA_MODEL))
        print("2: OpenAI (using {})".format(OPENAI_MODEL))
        choice = input("Enter your choice (1 or 2): ")
        if choice in ['1', '2']:
            return choice
        else:
            print("Invalid choice. Please enter 1 or 2.")

# -----------------------------
# 2) CHAT FUNCTION
# -----------------------------
def chat_with_model(user_input: str, system_prompt: str, api_choice: str, openai_client=None) -> str:
    """
    Sends user_input to the selected AI model (Ollama or OpenAI)
    using the provided system_prompt and returns the model's text output.
    """
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_input}
    ]

    try:
        model_name = OLLAMA_MODEL if api_choice == '1' else OPENAI_MODEL
        print(f"--- Calling { 'Ollama' if api_choice == '1' else 'OpenAI'} ({model_name}) ---")

        if api_choice == '1': # Ollama
            response = ollama.chat(model=OLLAMA_MODEL, messages=messages)
            return response["message"]["content"]

        elif api_choice == '2': # OpenAI
            if not openai_client:
                raise ValueError("OpenAI client not initialized.")
            response = openai_client.chat.completions.create(
                model=OPENAI_MODEL,
                messages=messages
            )
            return response.choices[0].message.content
        else:
             raise ValueError("Invalid API choice provided to chat function.") # Should not happen

    except Exception as e:
        print(f"\n--- ERROR during API call for an agent ---")
        # Check specifically for AuthenticationError if using OpenAI hardcoded key
        if isinstance(e, openai.AuthenticationError):
             print("Error: Authentication failed. Check if your hardcoded API key is correct and valid.")
        else:
            print(f"Error: {e}")
        print(f"System Prompt was: {system_prompt[:100]}...") # Print start of system prompt for context
        print(f"User Input was: {user_input}")
        print("Returning error message for this agent's response.")
        # Return a clear error message for the agent's specific output cell
        return f"ERROR: AI Agent call failed. Details: {e}"

# -----------------------------
# 4) MAIN PROCESS
# -----------------------------
def main():
    # Get user choice for API
    api_choice = get_api_choice()
    openai_client = None

    # Setup OpenAI client if chosen
    if api_choice == '2':
        hardcoded_api_key = "keyyyyyyyy"
        if not hardcoded_api_key or hardcoded_api_key == "PASTE_YOUR_OPENAI_API_KEY_HERE":
            print("\nERROR: You need to replace 'PASTE_YOUR_OPENAI_API_KEY_HERE' in the main() function")
            print("       with your actual OpenAI API key before running with option 2.")
            sys.exit(1)
        try:
            openai_client = openai.OpenAI(api_key=hardcoded_api_key)
            print("OpenAI client initialized using hardcoded key (INSECURE - FOR TESTING ONLY).")
        except openai.AuthenticationError:
            print("\nERROR: OpenAI Authentication Failed. Your hardcoded API key is likely incorrect or invalid.")
            sys.exit(1)
        except Exception as e:
            print(f"Error initializing OpenAI client: {e}")
            sys.exit(1)

    # Read the Excel file
    try:
        # Load the workbook and select the active worksheet
        workbook = load_workbook(filename='Multiple Claims.xlsx')
        sheet = workbook.active

        # Extract the data into a list of lists
        data = sheet.values

        # Get the column headers
        columns = next(data)

        # Create the DataFrame
        df = pd.DataFrame(data, columns=columns)
        print(df.head())
    except FileNotFoundError:
        print(f"ERROR: Input file '{INPUT_EXCEL_FILE}' not found.")
        sys.exit(1)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        sys.exit(1)

    # --- Ensure output columns exist for each agent ---
    print("Ensuring output columns exist for agents:")
    for agent_name in AGENT_PROMPTS.keys():
        col_name = agent_name.replace(" ", "_") + "_Response"
        if col_name not in df.columns:
            df[col_name] = ""
            print(f" - Added column: {col_name}")

    # Process each row
    total_rows = len(df)
    for idx, row in df.iterrows():
        print(f"\nProcessing Row {idx+1}/{total_rows}...")
        try:
            claim_id_original = str(row.get("Claim ID", f"Row_{idx+1}_NoID"))
            input_details = str(row.get("Input Details", ""))
            if not input_details or pd.isna(input_details) or input_details.lower() == 'nan':
                print(f"   Skipping row {idx+1} due to empty 'Input Details'.")
                continue
            user_prompt_base = f"Claim ID: {claim_id_original}\n{input_details}"
        except KeyError as e:
            print(f"   ERROR: Missing expected input column in row {idx+1}: {e}. Skipping row.")
            continue
        except Exception as e:
            print(f"   ERROR: Could not read required data for row {idx+1}: {e}. Skipping row.")
            continue

        # --- Loop through each defined Agent ---
        print(f"  Running Agents for Row {idx+1}:")
        for agent_name, system_prompt in AGENT_PROMPTS.items():
            print(f"   - Running Agent: {agent_name}")
            current_user_prompt = user_prompt_base
            agent_response = chat_with_model(
                user_input=current_user_prompt,
                system_prompt=system_prompt,
                api_choice=api_choice,
                openai_client=openai_client
            )
            col_name = agent_name.replace(" ", "_") + "_Response"
            df.at[idx, col_name] = agent_response.strip()
            print(f"     > Response stored in '{col_name}'.")
        print(f"  Finished all agents for Row {idx+1}.")

    # Save to a new file or overwrite
    try:
        df.to_excel(OUTPUT_EXCEL_FILE, index=False, engine='openpyxl')
        print(f"\nDone! Results saved to {OUTPUT_EXCEL_FILE}")
    except Exception as e:
        print(f"\nERROR saving results to Excel file: {e}")
        print("Please ensure the file is not open and you have write permissions.")


if __name__ == "__main__":
    # --- Prerequisites ---
    print("Agentic Medical Claim Analysis")
    print("=" * 30)
    print("Ensure required libraries are installed:")
    print("pip install pandas ollama openai openpyxl")
    print("-" * 30)
    # Added warning about hardcoded key if user chooses option 2
    print("NOTE: If you choose option 2 (OpenAI), you MUST edit the script")
    print("      and replace 'PASTE_YOUR_OPENAI_API_KEY_HERE' with your actual key.")
    print("      This is insecure and only intended for temporary testing.")
    print("-" * 30)
    main()