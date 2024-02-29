# AutoCAD Polyline Coordinate Exporter Plugin

Welcome to the repository for the AutoCAD Polyline Coordinate Exporter, a custom plugin designed to streamline and automate repetitive drafting tasks for AutoCAD users.

## Summary

This plugin was developed to enhance productivity in the design cycle, especially for projects that require the extraction and organization of polyline coordinates. With the use of C# and the .NET API for AutoCAD, this tool has been instrumental in boosting efficiency for over 50 distinct projects.

Key features of this plugin include:

- Automation of polyline coordinate extraction in AutoCAD.
- Collection and organization of geometry properties from more than 120 designed sections.
- Storage of polyline data in a well-structured and easily retrievable format using Microsoft Access.
- Use of ADO.NET technology to establish a robust connection between MS Access and the AutoCAD .NET API.
- A significant increase in drafting rate by 40% for designers using this plugin.

## Installation

To install the Polyline Coordinate Exporter plugin, follow these steps:

1. Download the plugin from the `Releases` section of this GitHub repository.
2. Open AutoCAD.
3. Use the `APPLOAD` command to load the plugin.
4. Browse to the location where you downloaded the plugin and select it.
5. Confirm that the plugin is loaded and will automatically load with future sessions of AutoCAD.

## Usage

After installing the plugin, you can begin exporting polyline coordinates with the following steps:

1. Open the AutoCAD drawing with the polylines you wish to export.
2. Run the command defined by the plugin (e.g., `EXPORTPLCOORDS`).
3. The plugin interface will guide you through selecting polylines and setting export options.
4. Choose the destination for the coordinate data, which will be organized and saved in an MS Access database.

## Requirements

- AutoCAD 2002 (replace with the specific versions supported)
- Microsoft Access
- .NET Framework (compatible with the version used by the plugin)

## Contributing

Contributions to the Polyline Coordinate Exporter plugin are welcome! To contribute, please fork the repository, make your changes, and submit a pull request for review.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

If you encounter any issues or have questions about using the plugin, please file an issue on this GitHub repository.

## Authors

- **[Xinrui Yi]** - *Initial work* - [Yillon Mask](https://github.com/YillonMask/)

We also acknowledge the contributions of all designers and developers who have used and improved this plugin.

## Acknowledgments

- Thank you to the AutoCAD .NET community for providing support and resources.
- Special thanks to my co-workers that have implemented this plugin, providing valuable feedback and suggestions.
