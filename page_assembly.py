from dataclasses import dataclass

import front_matter


@dataclass
class PageInsert:
    page_type: str
    text: str = ""
    image_path: str | None = None


def build_front_matter_pages(
    front_matter_data: front_matter.FrontMatter,
    paragraphs: list[str],
) -> tuple[list[PageInsert], set[int]]:
    pages = []
    used_paragraph_indexes = set()

    # Title Page
    title_page = front_matter_data.title_page

    if (
        title_page is not None
        and title_page.start_index is not None
        and title_page.end_index is not None
    ):
        text = "\n".join(
            paragraphs[
                title_page.start_index:
                title_page.end_index + 1
            ]
        )

        pages.append(
            PageInsert(
                text=text,
                page_type="title_page",
            )
        )

        used_paragraph_indexes.update(
            range(
                title_page.start_index,
                title_page.end_index + 1,
            )
        )

    # Copyright
    if front_matter_data.copyright is not None:
        pages.append(
            PageInsert(
                text=front_matter_data.copyright.text,
                page_type="copyright",
            )
        )

    # Dedication
    dedication = front_matter_data.dedication

    if (
        dedication is not None
        and dedication.start_index is not None
        and dedication.end_index is not None
    ):
        text = "\n".join(
            paragraphs[
                dedication.start_index:
                dedication.end_index + 1
            ]
        )

        pages.append(
            PageInsert(
                text=text,
                page_type="dedication",
            )
        )

        used_paragraph_indexes.update(
            range(
                dedication.start_index,
                dedication.end_index + 1,
            )
        )

    # Map
    if front_matter_data.map_file is not None:
        pages.append(
            PageInsert(
                page_type="map",
                image_path=front_matter_data.map_file,
            )
        )

    # Trigger Warnings
        if front_matter_data.trigger_warnings is not None:
            pages.append(
                PageInsert(
                    text=front_matter_data.trigger_warnings.text,
                    page_type="trigger_warnings",
                )
            )

    return pages, used_paragraph_indexes