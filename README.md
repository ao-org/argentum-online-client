## ⚔️ Support Us on Patreon

Please consider supporting our work on [Patreon](https://www.patreon.com/nolandstudios). Your support helps us continue developing and maintaining projects like this one. Every contribution makes a significant impact!

## 🐲 Argentum Online Client 🧙‍♂️

Welcome to the source code repository for the Argentum Online Client. To fully utilize this client, you will need the corresponding server and assets:

- **Server:** [Argentum Online Server](https://github.com/ao-org/argentum-online-server)
- **Assets:** [Resources](https://github.com/ao-org/Recursos)

## 🐛 Report bugs
- **Please report the bugs in the Server repository to maintain all of the tickets in one place:** [Link](https://github.com/ao-org/Recursos](https://github.com/ao-org/argentum-online-server/issues))

## 🛡️ Contributing via Pull Requests

We encourage contributions to the project! However, to maintain code quality and readability, please adhere to our guidelines, especially when dealing with the peculiarities of Visual Basic 6 IDE, which tends to change variable names, complicating the review process for Pull Requests.

### Setting Up Pre-commit Hooks

We utilize a pre-commit hook to minimize issues with variable name changes and other potential conflicts. Follow these steps to set up your environment correctly:

1. Open `git bash` or your preferred git client.
2. Execute the following commands:

```bash
chmod +x .githooks/pre-commit
git config core.hooksPath .githooks
```

Basically the pre-commit hook runs when you make a `git commit` and it will run the file `git_ignore_case.sh` to avoid false changes in the Pull Request. Is not perfect but it helps a lot. Please send the Pull Requests with only the neccesary code to be reviewed.

In case you have problems setting locally your pre-commit hook you can run the file `git_ignore_case.sh` by just doing double click.

<a href="https://imgbb.com/"><img src="https://i.ibb.co/6wCZvvZ/image.png" alt="Precommit-hook" border="0"></a>

This pre-commit hook executes the `git_ignore_case.sh` script during a `git commit`, helping to avoid false changes in Pull Requests. While it's not a perfect solution, it significantly aids in keeping our project clean and review-friendly.

### If You Encounter Issues

If you have any trouble setting up the pre-commit hook locally, you can manually run the `git_ignore_case.sh` script by double-clicking on it. This step ensures that your contributions are as clean and straightforward as possible.

![Contribution Guidelines](https://steamuserimages-a.akamaihd.net/ugc/1829034638748296385/CCD6BAF674692E8D4C87CDCA56FF8EC06D93C2FB/?imw=5000&imh=5000&ima=fit&impolicy=Letterbox&imcolor=%23000000&letterbox=false)

We appreciate your interest in contributing to the AO20 Client. By following these guidelines, you help us maintain a high standard of code quality and ensure that your contributions can be efficiently reviewed and integrated.

## Cryptography
CryptoSys is used in Argentum Online to cipher sensitive data.

- [https://www.cryptosys.net/api.html](https://www.cryptosys.net/api.html)

Please note this is not free software and you will have to buy your own license to use CryptoSys

## Star History

<a href="https://star-history.com/#ao-org/argentum-online-client&Date">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="https://api.star-history.com/svg?repos=ao-org/argentum-online-client&type=Date&theme=dark" />
    <source media="(prefers-color-scheme: light)" srcset="https://api.star-history.com/svg?repos=ao-org/argentum-online-client&type=Date" />
    <img alt="Star History Chart" src="https://api.star-history.com/svg?repos=ao-org/argentum-online-client&type=Date" />
  </picture>
</a>
