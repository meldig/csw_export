# csw_export

Le but de ce script est de récupérer toutes les fiches des différents catalogues ISOGEO de la MEL et de mettre en forme certaines informations dans des fichiers excel.
On va se servir du protocole CSW et de ses différentes méthodes.

## Script

Le fichier csw_export.py contient le code exécutable. On y retrouve plusieurs fonctions qui remplissent des tâches précises.

##### getRecords

La fonction getRecords récupère toutes les fiches des catalogues avec leurs informations. On réalise un appel à la méthode GetRecords (protocole CSW) et on récupère une réponse au format XML.
Cette réponse est traitée afin d'extraire les balises qui nous intéressent pour la mise en forme du fichier Excel, on réalise cette opération grâce à la méthode "findall" de la bibliothèque ElementTree.
Elle attend un paramètre : 
- url : l'url qui permettra de faire la requête "GetRecords"

##### writeInExcelFile

La fonction WriteInExcelFile permet de créer et d'écrire un fichier excel. Cette fonction attend trois paramètres : 
- records : tableau d'éléments provenant de la réponse XML de la méthode GetRecords du protocole CSW
- filePath : le chemin absolu du fichier excel que l'on veut créer
- url : l'url de base qui permet d'afficher une fiche en particulier sur le catalogue web
Dans cette fonction, on itère sur tous les éléments de "records" pour écrire dans le fichier excel.

## Configuration

Le fichier de configuration contient des urls, des chemins de fichiers et des balises XML que l'on utilise dans le script. Il est utile pour ne pas avoir à mettre ces informations en dur dans le code et donc améliorer la lisibilité.
Il est important de ne pas oublier la clé "type" dans chacun des blocs afin de bien identifier quel type d'information nous manipulons.
Le fichier de configuration est chargé en premier dans le fichier csw_export.py.

L'url qui sert à faire la requête GetRecords peut être trouvé grâce à la méthode GetCapabilities ("lien CSW" dans ISOGEO et balise "<ows:Operation>" avec le paramètre "name='GetRecords'", et on prendra le lien présent dans la balise "<ows:Post>")

## Fichier de requête

Le fichier request.xml contient le corps de la requête POST pour la méthode GetRecords du protocole CSW.
On y retrouve les différents paramètres acceptés par GetRecords comme outputFormat, resultType, typeNames, ElementSetName etc...

## Remarque

Le fichier main et le fichier de configuration présents sur le repo sont des fichiers génériques à modifier avec des valeurs personnelles.