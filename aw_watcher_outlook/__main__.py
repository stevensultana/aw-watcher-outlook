from aw_watcher_outlook.outlook import main
from aw_watcher_outlook.config import parse_args

if __name__ == "__main__":
    args = parse_args()
    main(args.poll_time, args.testing)
