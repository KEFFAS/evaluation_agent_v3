from modules.ce.workflow import run_ce

report = run_ce(
    excel_file=r"uploads\SSDC 6 CE.xlsx",

    programme_title="Your Programme",

    programme_code="ABC 01/2026",

    duration="1st July 2026 to 5th July 2026",

    venue="KSG Matuga",

    coordinator="John Doe",

    assistant="Jane Doe",

    generate_analysis=True,

    output_folder="outputs"
)

print(report)