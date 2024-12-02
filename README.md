# Modelleren van oppervlaktewater
Deze repository bevat al het cursusmateriaal van de casus 'Landelijk Water' binnen de major Watermanagement van Hogeschool Van Hall Larenstein.

De map Syllabus bevat de broncode van de opdrachten. We beheren de opdrachten in [Quarto](https://quarto.org/docs/download/).
Quarto is een publicatieplatform waarmee dynamische content geproduceerd en gepubliceerd kan worden, in combinatie met bijvoorbeeld R, Python, Julia etc.

De opdrachten worden door Quarto gerenderd naar HTML/CSS/Javascript, in de map docs.
De inhoud van docs is een website die op alle webservers gehost kan worden, maar die ook gewoon lokaal kan worden geopend, door op index.html te klikken.

Github biedt een gratis mogelijkheid om deze documentatie uit docs te hosten. Dit hebben we als zodanig ingesteld onder 'pages'. 
Het gehoste cursusmateriaal kan [hier](https://siebebosch.github.io/modelleren_oppervlaktewater/) worden gevonden.

## Installeren
Om aan het cursusmateriaal te kunnen werken is het volgende nodig:

* Een installatie van Microsoft [Visual Studio Code](https://code.visualstudio.com/download)
* Een installatie van [Quarto](https://quarto.org/docs/download/)
* Een installatie van [Git](https://gitforwindows.org/)

Als deze pakketten geïnstalleerd zijn kunt u de de volgende extensies installeren in Visual Studio Code:
* Quarto 
* Github Actions

Vergeet niet je github-account te configureren in de laatstgenoemde extension.

## De repository beheren

Het beheren van de repository kan op verschillende manieren. Bijvoorbeeld het programma Github for Desktop is een populaire methode. 
Hier geven we echter een stappenplan, uitgaande van werken vanaf de command line.

### Initialiseren
blader naar de projectmap
```cd your-project-folder```

de git repository initialiseren:
```git init```

de git repository toevoegen
```git remote add origin https://github.com/SiebeBosch/modelleren_oppervlaktewater.git```

een lokale kopie van de repository naar je eigen computer halen
```git pull origin main```

### Wijzigingen committen en pushen

alle nieuwe bestanden indexeren
```git add . ```

aanpassingen committen
```git commit -m "omschrijving"```

je aanpassingen committen naar de repository
```git push -u origin main```

### Wijzigingen in je lokale repository synchroniseren met de online repository

Als anderen wijzigingen hebben aangebracht in de repository is het van belang om die Wijzigingen ook door te voeren in je lokale kopie van de repository. Dit gebeurt met het commando 'pull'.

```git pull```
