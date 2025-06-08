"""get latest release version via  GET /repos/{owner}/{repo}/releases/latest
check in local, if local version < remote version, then call GET /repos/{owner}/{repo}/releases/{release_id}/assets
to get all the assets of the latest release
Get the id of the one you determine that is the new installer, download it via
GET /repos/{owner}/{repo}/releases/assets/{asset_id}
make sure you tweak your installer to check existing directory, if it is already exist, then must keep the properties
files and the logs also

Need to store a file in the /input to help retrieving the local version
store your access token and the GitHub endpoints in a separated json file in the source"""
import os
import re
import subprocess
import time

import requests

from src.common.FileUtil import load_key_value_from_file_properties, get_content_of_a_file_as_a_line
from src.common.ProcessUtil import get_matching_processes, kill_processes
from src.setup.packaging.path.PathResolvingService import PathResolvingService
from src.setup.update.model.GetReleaseResponse import Release, Asset

config_file_path: str = os.path.join(PathResolvingService.get_instance().get_input_dir(),
                                     'Version_Control_Getting_Release.properties')
api_getting_release_client_settings: dict[str, str] = load_key_value_from_file_properties(config_file_path)


def update_on_demand():
    local_release: Release = get_local_running_release()
    remote_release: Release = get_latest_remote_release()

    local_version = validate_and_extract_version(local_release.tag_name)
    remote_version = validate_and_extract_version(remote_release.tag_name)

    if is_needed_update(local_version, remote_version) is False:
        # log here, or notify UI already latest version
        return

    installer_full_path: str = download_asset(remote_release)
    process = subprocess.Popen([installer_full_path, '/NOSCREENS'],
                               stdout=subprocess.PIPE,
                               stderr=subprocess.PIPE,
                               text=True)

    while len(get_matching_processes('Setup - automation_tool')) > 0:
        time.sleep(1)

    kill_processes('automation_tool')
    # process = subprocess.Popen([installer_full_path, '/NOSCREENS'],
    #                            stdout=subprocess.PIPE,
    #                            stderr=subprocess.PIPE,
    #                            text=True)


def download_asset(remote_release: Release) -> str:
    download_url = remote_release.assets[0].url
    storing_dir = PathResolvingService.get_instance().get_release_notes()
    installer_path = os.path.join(storing_dir, 'installer.exe')

    headers = {
        "Accept": "application/octet-stream",
        "X-GitHub-Api-Version": api_getting_release_client_settings['github.release.version'],
        "Authorization": 'Bearer ' + api_getting_release_client_settings['github.pat']
    }
    response = requests.get(download_url, headers=headers, stream=True)
    response.raise_for_status()

    with open(installer_path, 'wb') as f:
        for chunk in response.iter_content(chunk_size=8192):
            f.write(chunk)
    return installer_path


def is_needed_update(local_version: list[int], remote_version: list[int]) -> bool:
    if len(local_version) != len(remote_version):
        raise ValueError("Local and remote version lists must have equal length")

    for local, remote in zip(local_version, remote_version):
        if local < remote:
            return True

    return False


def validate_and_extract_version(tag_name: str) -> list[int]:
    pattern = r'^v(\d{1,2})\.(\d{1,2})\.(\d{1,2})$'

    # Match the input string against the pattern
    match = re.match(pattern=pattern, string=tag_name)

    if match:
        parts = [int(match.group(1)), int(match.group(2)), int(match.group(3))]
        return parts
    else:
        raise Exception('The tag version is not in format {}', pattern)


def get_local_running_release() -> Release:
    current_local_release = os.path.join(PathResolvingService.get_instance().get_release_notes(),
                                         'current_release_version.txt')
    content: str = get_content_of_a_file_as_a_line(current_local_release)
    return Release(tag_name=content)


def get_latest_remote_release() -> Release:
    url = api_getting_release_client_settings['github.release.latest.url']
    headers = {
        "X-GitHub-Api-Version": api_getting_release_client_settings['github.release.version'],
        "Authorization": 'Bearer ' + api_getting_release_client_settings['github.pat']
    }

    response = requests.get(url, headers=headers)
    response.raise_for_status()
    data = response.json()

    assets = [
        Asset(
            url=asset["url"],
            id=asset["id"],
            node_id=asset["node_id"],
            name=asset["name"],
            label=asset["label"],
            content_type=asset["content_type"],
            state=asset["state"],
            size=asset["size"],
            digest=asset["digest"],
            download_count=asset["download_count"],
            created_at=asset["created_at"],
            updated_at=asset["updated_at"],
            browser_download_url=asset["browser_download_url"]
        )
        for asset in data["assets"]
    ]

    release = Release(
        url=data["url"],
        assets_url=data["assets_url"],
        upload_url=data["upload_url"],
        html_url=data["html_url"],
        id=data["id"],
        node_id=data["node_id"],
        tag_name=data["tag_name"],
        target_commitish=data["target_commitish"],
        name=data["name"],
        draft=data["draft"],
        prerelease=data["prerelease"],
        created_at=data["created_at"],
        published_at=data["published_at"],
        assets=assets,
        tarball_url=data["tarball_url"],
        zipball_url=data["zipball_url"],
        body=data["body"]
    )

    return release
