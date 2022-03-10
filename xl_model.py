import dataclasses


@dataclasses.dataclass
class ExcelModel:
    # All Rows from left to right in Excel starting at 1 (0), data starts at Col 10
    Unternehmen: str
    Geschaeftsfuehrer: str
    Telefon: str
    eMail: str
    Adresse: str
    MoritzCheck: str
    Anschreiben: str
    AnfrageLinkedin: str
    Abgeschickt: str
    Interessant: str
    Mail: str
    Link: str  # Col 15


@dataclasses.dataclass
class ExToCompare:
    eMail: str
    Unternehmen: str
    Telefon: str
