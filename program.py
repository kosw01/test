from CablebridgeAnalysis import BridgeAnalysis

br_name = '화태대교'
bridge = BridgeAnalysis(br_name)
bridge.plot_time_history()
bridge.plot_scatter()
bridge.generate_summary_report()
bridge.generate_summary_report_excel()
# bridge.generate_summary_report_word()
# bridge.calculate_weekly_reception_rate()
