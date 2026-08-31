import constants

def calculate_gutter_width(page_count: int) -> float:
    if page_count <= 150:
        return 0.375
    elif page_count <= 300:
        return 0.500
    elif page_count <= 500:
        return 0.625
    elif page_count <= 700:
        return 0.750
    else:
        return 0.875

def inches_to_pixels(inches: float) -> int:
    return round(inches * constants.PIXELS_PER_INCH)