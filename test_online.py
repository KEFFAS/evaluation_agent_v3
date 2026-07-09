from modules.online.workflow import run_online

report = run_online(

    csv_file=r"uploads\PRCC 14.csv",

    programme_title="Senior Management Course",

    programme_code="PRCC 14/2026",

    duration="14th July 2026 to 18th July 2026",

    venue="KSG Online",

    coordinator="John Doe",

    assistant="Jane Doe",

    output_folder="outputs"

)

print("\nGenerated Report:")
print(report)