# PnP Sites Core Library

The [PnP Sites Core library](https://github.com/PnP/PnP-Sites-Core) is very popular library that extends SharePoint using mainly CSOM. This library contains the PnP Provisioning engine, tons of extension methods, a modern page API, etc...but this library has also organically grown into a complex and hard to maintain code base. One of the reasons why the [PnP Core SDK](https://github.com/pnp/pnpcore) development started is to provide a new clean basis for the PnP Sites Core library with a strong focus on quality (test coverage above 80%, automation). As this transition will take quite some time and effort we plan to gradually move things over from PnP Sites Core to the PnP Core SDK. The first step in this transition is releasing a .Net Standard 2.0 version of PnP Sites Core, called [PnP Framework](https://github.com/pnp/pnpframework). Going forward [PnP Framework](https://github.com/pnp/pnpframework) features will move to the PnP Core SDK in a phased approach. At this moment we've shipped our first [PnP Framework](https://github.com/pnp/pnpframework) preview version and preview 3 of the [PnP Core SDK](https://github.com/pnp/pnpcore). 

> **Important:**
> PnP Sites Core will be retired by the end of 2020. As of the GA of [PnP Framework](https://github.com/pnp/pnpframework) we'll only maintain that version going forward.

![PnP dotnet roadmap](PnP%20dotnet%20Roadmap%20-%20October%20status.png)

## I've found a bug, where do I need to log an issue or create a PR

Between now and the end of 2020 both [PnP-Sites-Core](https://github.com/PnP/PnP-Sites-Core) and [PnP Framework](https://github.com/pnp/pnpframework) are actively maintained. Once [PnP Framework](https://github.com/pnp/pnpframework) GA's we'll stop maintaining [PnP-Sites-Core](https://github.com/PnP/PnP-Sites-Core).

Given [PnP Framework](https://github.com/pnp/pnpframework) is our future going forward we would prefer issues and PR's being created in the [PnP Framework](https://github.com/pnp/pnpframework) repo. If you want your PR to apply to both then it's recommended to create the PR in both repositories for the time being.

**Community rocks, sharing is caring!**

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.