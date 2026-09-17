from create_contest_proposal import Image, main


_original_new = Image.new


def _new_with_hex_compat(mode, size, color=0):
    if isinstance(color, str) and len(color) in (6, 8) and not color.startswith("#"):
        color = "#" + color
    return _original_new(mode, size, color)


Image.new = _new_with_hex_compat


if __name__ == "__main__":
    main()
