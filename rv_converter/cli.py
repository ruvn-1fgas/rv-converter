import argparse
import sys

from .core import converter
from .i18n import get_i18n
from .utils import load_data

i18n = get_i18n()


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=i18n["program_description"])
    parser.add_argument("filename", help=i18n["filename_help"])
    parser.add_argument(
        "-o",
        "--output",
        help=i18n["output_help"],
    )
    parser.add_argument(
        "-s",
        "--sort-by",
        dest="sort_by",
        help=i18n["sort_by_help"],
    )
    parser.add_argument(
        "-j",
        "--join",
        dest="join_multivalued",
        help=i18n["join_multivalued_help"],
    )
    parser.add_argument(
        "-t",
        "--type",
        help=i18n["type_help"],
        choices=["xlsx", "csv", "json", "jsonl"],
        default="xlsx",
    )

    return parser.parse_args()


def main() -> None:
    args = parse_arguments()
    data = load_data(args.filename)

    if data is None:
        print(i18n["error_data_none"])
        sys.exit(1)

    if not data:
        print(i18n["error_data_empty"])
        sys.exit(1)

    if not isinstance(data[0], dict):
        print(i18n["error_data_format"])
        sys.exit(1)

    output_type = args.type
    output_file = (
        args.output.rsplit(".", 1)[0] + f".{output_type}"
        if args.output
        else args.filename.rsplit(".", 1)[0] + f".{output_type}"
    )

    if output_type in ("json", "jsonl"):
        converter.save_file(
            data,
            output_file,
            sort_by=args.sort_by,
            join_multivalued=args.join_multivalued,
            type=output_type,
        )
    else:
        converter.save_file(
            data,
            output_file,
            sort_by=args.sort_by,
            join_multivalued=args.join_multivalued,
            type=output_type,
        )


if __name__ == "__main__":
    main()
