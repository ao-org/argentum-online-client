#### ⚔️ Por favor considera apoyarnos en [https://www.patreon.com/nolandstudios](https://www.patreon.com/nolandstudios) ⚔️

# 🐲 AO20 CLIENTE 🧙🏻
Código fuente del cliente de Argentum20.

Para utilizar este cliente, necesitas el servidor correspondiente y assets:
[argentum20-server](https://github.com/ao-org/argentum20-server)
[Recursos](https://github.com/ao-org/Recursos)


## Por favor considera apoyarnos en [Patreon](https://www.patreon.com/nolandstudios)

# 🛡️ Pull Requests

<a href="https://imgbb.com/"><img src="https://i.ibb.co/6wCZvvZ/image.png" alt="Precommit-hook" border="0"></a>

We have a pre-commit hook for the project, Visual Basic 6 IDE changes the names of the variables and it makes the Pull Requests very difficult to understand.

Please run the following commands with `git bash` or the client you are using.

```
chmod +x .githooks/pre-commit
git config core.hooksPath .githooks
```

Basically the pre-commit hook runs when you make a `git commit` and it will run the file `git_ignore_case.sh` to avoid false changes in the Pull Request. Is not perfect but it helps a lot. Please send the Pull Requests with only the neccesary code to be reviewed.

In case you have problems setting locally your pre-commit hook you can run the file `git_ignore_case.sh` by just doing double click.


![PR Image](https://steamuserimages-a.akamaihd.net/ugc/1829034638748296385/CCD6BAF674692E8D4C87CDCA56FF8EC06D93C2FB/?imw=5000&imh=5000&ima=fit&impolicy=Letterbox&imcolor=%23000000&letterbox=false)

