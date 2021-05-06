# Exercices de révision sur les tests unitaires
Voici quelques exercices qui vont vous aider pour l'éventuel évaluation sur ce sujet.

## Scénario
Vous êtes attitré à un projet de développement d'un système de récensement des espèces écologiques. Le système doit extraire les données d'un fichier Excel source vers un autre fichier Excel. Le projet C# est déjà débuté et il utilise le package Nuget ClosedXML pour gérer les fichiers Excel.

## Requis du domaine
- Les contenus des cellules A1, B1, C1 et D1 de la première feuille seront respectivement _record_id, _title, _server_updated_at et _updated_by
- Le fichier Excel devra avoir au moins 5 feuilles. Sinon un message à l'utilisateur indiquera que le fichier source n'est pas valide.

## Requis applicatifs
L'application doit être en mesure de charger un fichier Excel.
**Un fichier Excel pour être valide**
- Le fichier doit exister. Sinon afficher un message à l'utilisateur indiquant que le fichier est inexistant.
- Le fichier Excel __source__ ne doit pas être en cours d'utilisation par Excel. Si tel est le cas, l'application ouvrira le fichier en lecture seule.
