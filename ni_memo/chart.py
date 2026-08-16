"""Small deterministic charts for traceable workbook year series."""
from io import BytesIO
import re

from PIL import Image, ImageDraw


YEAR = re.compile(r"^(?:19|20)\d{2}[AE]$", re.IGNORECASE)
COLORS = ((24, 137, 142), (31, 77, 120), (211, 132, 57), (104, 91, 160))


def render_chart_png(region):
    year_columns = [
        (index, str(value).strip().upper())
        for index, value in enumerate(region.headers)
        if value is not None and YEAR.fullmatch(str(value).strip())
    ]
    if len(year_columns) < 2:
        return None
    series = []
    for row in region.rows[1:]:
        values = [row[index] if index < len(row) else None for index, _ in year_columns]
        if all(_number(value) is not None for value in values):
            label = next((str(value) for value in row[:year_columns[0][0]] if value not in (None, "")), "Series")
            series.append((label, tuple(float(value) for value in values)))
        if len(series) == len(COLORS):
            break
    if not series:
        return None
    return _draw(tuple(year for _, year in year_columns), series)


def render_series_chart_png(periods, series):
    """Render already-frozen semantic fact series without a workbook-region dump."""
    periods = tuple(str(value).strip().upper() for value in periods)
    series = tuple((str(label), tuple(float(value) for value in values))
                   for label, values in series)
    if len(periods) < 2 or not series:
        return None
    if any(len(values) != len(periods) for _, values in series):
        return None
    return _draw(periods, series[:len(COLORS)])


def _draw(years, series):
    width, height = 1200, 620
    left, top, right, bottom = 105, 55, 50, 95
    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    plot_width, plot_height = width - left - right, height - top - bottom
    values = [value for _, row in series for value in row]
    low, high = min(values), max(values)
    low = min(0.0, low)
    if high == low:
        high = low + 1.0
    draw.rectangle((left, top, left + plot_width, top + plot_height), outline=(190, 198, 205), width=2)
    for step in range(5):
        y = top + plot_height * step / 4
        draw.line((left, y, left + plot_width, y), fill=(229, 233, 237), width=1)
        label = high - (high - low) * step / 4
        draw.text((10, y - 7), f"{label:,.1f}", fill=(85, 94, 103))
    x_positions = [left + plot_width * index / (len(years) - 1) for index in range(len(years))]
    for x, year in zip(x_positions, years):
        draw.text((x - 18, top + plot_height + 18), year, fill=(65, 73, 82))
    for series_index, (label, row) in enumerate(series):
        color = COLORS[series_index]
        points = [(x, top + (high - value) / (high - low) * plot_height) for x, value in zip(x_positions, row)]
        draw.line(points, fill=color, width=5)
        for x, y in points:
            draw.ellipse((x - 6, y - 6, x + 6, y + 6), fill=color)
        draw.line((left + series_index * 130, height - 28, left + series_index * 130 + 30, height - 28),
                  fill=color, width=5)
        legend = str(label).strip()[:12] or f"Series {series_index + 1}"
        draw.text((left + series_index * 130 + 38, height - 36), legend, fill=(65, 73, 82))
    output = BytesIO()
    image.save(output, format="PNG", optimize=False)
    return output.getvalue()


def _number(value):
    return value if isinstance(value, (int, float)) and not isinstance(value, bool) else None
