### Structure du classeur

Le classeur de NotaComp est composé de 2 feuilles fixes, suivies de 2 feuilles par classe générée. Chaque Page est gérée par le module du même index (exemple: Module3 gère la Page 3).

- Page 1 (```Accueil```): L'utilisateur entre ici ses caractéristiques (nom, matière, établissement, année scolaire) ainsi que les paramètres généraux de NotaComp (nombre de domaines/compétences, nombre de classes/élèves).
- Page 2 (```Liste de classes```): L'utilisateur entre ici le nom des élèves de chaque classe qu'il a précédemment déclaré.
- Page 3 (```Notes```): L'utilisateur entre ici les notes obtenues par les élèves pour chaque évaluation. Il peut ajouter/modifier/supprimer des évaluations selon ses besoins. Il existe une feuille par classe.
- Page 4 (```Bilan```): L'utilisateur consulte ici le bilan semestriel et annuel par domaine, ainsi que la moyenne trimestrielle et annuelle. Il existe une feuille par classe.


### Structure du programme

Le programme de NotaComp est divisé en 4 modules, chacun contenant
des procédures et fonction permettant d'interagir avec les feuilles
du classeur. Chaque module est dédié à une ou plusieurs feuilles
spécfifiques:
* Module1 - Gère la Page1 "Accueil".
 Fonction unique: génère la Page2 "Liste de classes".
* Module2 - Gère la Page2 "Liste". Ce module s'interface avec
 les UserForm 1 à 5 pour effectuer les opérations de modification.
 Fonction unique: génère les Page3 "Notes" et les Page4 "Bilan".
* Module3 - Gère la Page3 "Notes".
* Module4 - Gère la Page4 "Bilan" en récupérant les données entrées
 dans les évaluations Page3.
* Module5 - Permet d'exporter les Modules et UserForms pour sauvegarde,
 et/ou de les importer dans un nouveau classeur.

*******************************************************************************

