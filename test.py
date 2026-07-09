from modules.eee.workflow import run_eee

report = run_eee(
    excel_file=r"E:\evaluation_agent_v3\uploads\PRCC 19 EE.xlsx",
    duration="3rd May 2026 to 10th May 2026",
    use_llm=True
)

print(report)