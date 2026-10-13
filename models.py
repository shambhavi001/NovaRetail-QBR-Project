from pydantic import BaseModel


class DeckContent(BaseModel):

    summary: str

    performance_bullets: list[str]

    customer_bullets: list[str]

    risk_bullets: list[str]

    recommendations: list[str]

    closing_statement: str

    slide2_headline: str
    slide3_headline: str
    slide4_headline: str

