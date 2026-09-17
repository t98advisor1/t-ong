from PIL import ImageColor
from create_contest_proposal import main


_original_getrgb = ImageColor.getrgb


def _getrgb_with_hex_compat(color):
    if isinstance(color, str) and len(color) in (6, 8) and not color.startswith("#"):
        color = "#" + color
    return _original_getrgb(color)


ImageColor.getrgb = _getrgb_with_hex_compat


if __name__ == "__main__":
    main()
