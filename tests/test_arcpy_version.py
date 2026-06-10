import arcpy


def get_arcpy_version_line() -> str:
    install_info = arcpy.GetInstallInfo()
    version = install_info.get("Version", "unknown")
    product_name = install_info.get("ProductName", "ArcGIS")
    return f"ArcPy version: {version} ({product_name})"


def test_print_arcpy_version():
    version_line = get_arcpy_version_line()
    print(version_line)
    assert "unknown" not in version_line


if __name__ == "__main__":
    print(get_arcpy_version_line())
