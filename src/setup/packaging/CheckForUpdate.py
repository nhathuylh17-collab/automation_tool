"""get latest release version via  GET /repos/{owner}/{repo}/releases/latest
check in local, if local version < remote version, then call GET /repos/{owner}/{repo}/releases/{release_id}/assets
to get all the assets of the latest release
Get the id of the one you determine that is the new installer, download it via
GET /repos/{owner}/{repo}/releases/assets/{asset_id}
make sure you tweak your installer to check existing directory, if it is already exist, then must keep the properties
files and the logs also

Need to store a file in the /input to help retrieving the local version
store your access token and the github etndpoints in a separated json file in the source"""
