from modules.fe.workflow import run_fe

report = run_fe(
    excel_file=r"uploads\SSDC 60 FE.xlsx",

    programme_title="Supervisory Skills Development Course",

    programme_code="SSDC 60/2026",

    duration="5 Days",

    venue="KSG Matuga",

    coordinator="John Doe",

    assistant="Jane Doe",

    total_participants=72,

    use_llm=True,

    generate_analysis=True,

    output_folder="outputs"
)

print("\nGenerated Report:")
print(report)