# TestTeam_Activity
**Prérequis** : un Python avec openpyxl et matplotlib installés

**Fonctionnement** : 
Le fichier excel ne doit contenir que les feuilles d'activités devant être contenues dans le graph

* Supprimer les feuilles inutiles (semaines ne devant pas apparaitre/feuilles ne contenant pas d'activités (y compris les feuilles cachées)
Au besoin éditer le fichier graph_activites.py et régler la variable xlsx_file sur le nom du fichier excel.
* Lancer le fichier python (commande python graph_activites.py)	
* Un graph cumulé est affiché et peut être sauvegardé. Un fichier Output.xlsx est généré et contient toutes les données dans la même feuille
