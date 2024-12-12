class API_LINKS:
    """Ссылки на данные в towerstore в json формате."""
    BASE_LINK: str = r'https://fridge.newtowers.ru/api/report'
    TL: str = f'{BASE_LINK}/sitesTable/json'
    KITS: str = f'{BASE_LINK}/placing/kits/json'
    TELE_2: str = f'{BASE_LINK}/placing/assets/tele2/json'
    MTS: str = f'{BASE_LINK}/placing/assets/mts/json'
    OTHER: str = f'{BASE_LINK}/placing/assets/other/json'
    MEGAFON: str = f'{BASE_LINK}/placing/assets/megafon/json'
    VIMPELCOM: str = f'{BASE_LINK}/placing/assets/vimpelcom/json'
    CONTRACT: str = f'{BASE_LINK}/tl/Table/json'


api_links = API_LINKS
