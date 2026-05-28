from workforce_allocation.pipeline import run_pipeline


def main() -> None:
    result = run_pipeline()
    print("TARGET_DATE:", result.target_date)
    print("Solver Status:", result.solver_status)
    print("Moved count:", result.moved_count)
    print("Exported:", result.output_file)


if __name__ == "__main__":
    main()

