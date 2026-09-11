from pathlib import Path


SUPPORTED_EXTENSIONS = {
    ".xlsx",
    ".csv",
}


def get_supported_files(folder):
    folder = Path(folder)

    if not folder.exists():
        raise FileNotFoundError(
            f"Folder does not exist: {folder}"
        )

    if not folder.is_dir():
        raise NotADirectoryError(
            f"Not a folder: {folder}"
        )

    return [
        file_path
        for file_path in folder.iterdir()
        if file_path.is_file()
        and file_path.suffix.lower()
        in SUPPORTED_EXTENSIONS
    ]


def normalize_filename(file_path):
    name = Path(file_path).stem

    return "".join(
        name.split()
    ).lower()


def find_pairs_by_filename(
    old_files,
    new_files,
):
    candidate_map = {}

    for old_index, old_file in enumerate(old_files):
        old_name = normalize_filename(old_file)

        candidate_map[old_index] = [
            new_index
            for new_index, new_file in enumerate(new_files)
            if (
                old_file.suffix.lower()
                == new_file.suffix.lower()
                and old_name
                in normalize_filename(new_file)
            )
        ]

    matched_new_to_old = {}

    def find_match(
        old_index,
        visited,
    ):
        candidates = candidate_map.get(
            old_index,
            [],
        )

        candidates = sorted(
            candidates,
            key=lambda new_index: (
                abs(
                    len(
                        normalize_filename(
                            new_files[new_index]
                        )
                    )
                    - len(
                        normalize_filename(
                            old_files[old_index]
                        )
                    )
                ),
                new_index,
            ),
        )

        for new_index in candidates:
            if new_index in visited:
                continue

            visited.add(new_index)

            previous_old_index = matched_new_to_old.get(
                new_index
            )

            if (
                previous_old_index is None
                or find_match(
                    previous_old_index,
                    visited,
                )
            ):
                matched_new_to_old[new_index] = old_index
                return True

        return False

    for old_index in range(len(old_files)):
        find_match(
            old_index,
            set(),
        )

    matched_old_indices = set(
        matched_new_to_old.values()
    )

    pairs_by_old_index = []

    for new_index, old_index in matched_new_to_old.items():
        pairs_by_old_index.append(
            (
                old_index,
                new_index,
            )
        )

    pairs_by_old_index.sort(
        key=lambda item: item[0]
    )

    pairs = [
        (
            old_files[old_index],
            new_files[new_index],
        )
        for old_index, new_index
        in pairs_by_old_index
    ]

    unmatched_old = [
        old_file
        for old_index, old_file
        in enumerate(old_files)
        if old_index not in matched_old_indices
    ]

    matched_new_indices = set(
        matched_new_to_old.keys()
    )

    unmatched_new = [
        new_file
        for new_index, new_file
        in enumerate(new_files)
        if new_index not in matched_new_indices
    ]

    return (
        pairs,
        unmatched_old,
        unmatched_new,
    )


def find_pairs_by_order(
    old_files,
    new_files,
):
    pairs = []
    unmatched_old = []
    unmatched_new = []

    pair_count = min(
        len(old_files),
        len(new_files),
    )

    for index in range(pair_count):
        old_file = old_files[index]
        new_file = new_files[index]

        if (
            old_file.suffix.lower()
            != new_file.suffix.lower()
        ):
            unmatched_old.append(
                old_file
            )
            unmatched_new.append(
                new_file
            )
            continue

        pairs.append(
            (
                old_file,
                new_file,
            )
        )

    if len(old_files) > pair_count:
        unmatched_old.extend(
            old_files[pair_count:]
        )

    if len(new_files) > pair_count:
        unmatched_new.extend(
            new_files[pair_count:]
        )

    return (
        pairs,
        unmatched_old,
        unmatched_new,
    )


def find_pairs(
    old_folder,
    new_folder,
    pairing_method,
):
    old_files = get_supported_files(
        old_folder
    )

    new_files = get_supported_files(
        new_folder
    )

    if pairing_method == "filename":
        return find_pairs_by_filename(
            old_files,
            new_files,
        )

    if pairing_method == "order":
        return find_pairs_by_order(
            old_files,
            new_files,
        )

    raise ValueError(
        f"Unknown pairing method: {pairing_method}"
    )
