from dataclasses import dataclass, field
from typing import List


@dataclass
class Asset:
    url: str
    id: int
    node_id: str
    name: str
    label: str
    content_type: str
    state: str
    size: int
    digest: str
    download_count: int
    created_at: str
    updated_at: str
    browser_download_url: str


@dataclass
class Release:
    tag_name: str
    url: str = field(default="", init=True)
    assets_url: str = field(default="", init=True)
    upload_url: str = field(default="", init=True)
    html_url: str = field(default="", init=True)
    id: int = field(default=None, init=True)
    node_id: str = field(default="", init=True)
    target_commitish: str = field(default="", init=True)
    name: str = field(default="", init=True)
    draft: bool = field(default=False, init=True)
    prerelease: bool = field(default=False, init=True)
    created_at: str = field(default="", init=True)
    published_at: str = field(default="", init=True)
    assets: List[Asset] = field(default=None, init=True)
    tarball_url: str = field(default="", init=True)
    zipball_url: str = field(default="", init=True)
    body: str = field(default="", init=True)
