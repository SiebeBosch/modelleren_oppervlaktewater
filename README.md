# Modelleren van oppervlaktewater
Deze repository bevat al het cursusmateriaal van de casus 'Landelijk Water' binnen de major Watermanagement van Hogeschool Van Hall Larenstein.

De map Syllabus bevat de broncode van de opdrachten. We beheren de opdrachten in [Quarto](https://quarto.org/docs/download/).
Quarto is een publicatieplatform waarmee dynamische content geproduceerd en gepubliceerd kan worden, in combinatie met bijvoorbeeld R, Python, Julia etc.

De opdrachten worden door Quarto gerenderd naar HTML/CSS/Javascript, in de map docs.
De inhoud van docs is een website die op alle webservers gehost kan worden, maar die ook gewoon lokaal kan worden geopend, door op index.html te klikken.

Github biedt een gratis mogelijkheid om deze documentatie uit docs te hosten. Dit hebben we als zodanig ingesteld onder 'Settings' - 'pages'. 
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
blader naar de gewenste projectmap
```cd your-project-folder```

de online git repository klonen:
```git clone https://github.com/SiebeBosch/modelleren_oppervlaktewater.git```

### Wijzigingen committen en pushen
alle nieuwe bestanden indexeren
```git add . ```

aanpassingen committen
```git commit -m "omschrijving"```

je aanpassingen pushen naar de repository
```git push -u origin main```

### Wijzigingen in je lokale repository synchroniseren met de online repository

Als anderen wijzigingen hebben aangebracht in de repository is het van belang om die Wijzigingen ook door te voeren in je lokale kopie van de repository. Dit gebeurt met het commando 'pull'.

```git pull```

## Schrijven
Quarto heeft een aantal basisstructuren bij het schrijven van teksten. Zo zijn er secties waarin je waarschuwingen, opmerkingen of vragen kunt formuleren:

```
::: {.callout-note}
### Informatie
Hier komt uw uitleg...
:::
```

```
::: {.callout-tip}
### Vraag
Hier komt uw vraag...
:::
```

```
::: {.callout-important}
### Belangrijk
Hier komt een waarschuwing...
:::
```

```
::: {.callout-warning}
### Let op!
Hier komt een waarschuwing...
:::
```

```
::: {.callout-caution}
### Voorzichting
Hier komt een waarschuwing level 2...
:::
```

Verder kan het gebruik van kolommen nuttig zijn:


```
::: {.columns}
::: {.column width="50%"}
De modelschematisatie kan automatisch worden vervaardigd met de bijbehorende bronbestanden en configuraties. Zo vervaardigen we het boezemmodel met de programma's Channel Builder en Catchment Builder, versie 3.0.0.2. (Hydroconsult).

:::
::: {.column width="50%"}
![Boezemmodel Noorderzijlvest in SOBEK](img/boezemmodel_sobek.png)
:::
:::
```

Afbeeldingen invoegen kan als volgt:
```
![Boezemmodel Noorderzijlvest in SOBEK](img/boezemmodel_sobek.png)
```
Bullets kunnen genummerd of met een asterisk.

```
* dit is een item
* dit is nog een item
```
```
1. Dit is item 1
2. Dit is item 2
```


