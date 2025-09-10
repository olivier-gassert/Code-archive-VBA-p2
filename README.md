# Code-archive-VBA-Partie-2

Ce projet illustre mes premiers pas dans le développement avec VBA, réalisés dans le but de digitaliser et réorganiser l’administration d’une boutique.

## Partie 2 : Programme de facturation

Cette seconde étape m’a permis de créer un **premier programme réellement utile** avec le langage VBA. Vers 2007–2008, les solutions logicielles étaient bien plus coûteuses, et la boutique ne pouvait pas se permettre un programme "tout-en-un". J’ai donc développé mes propres outils sur Excel afin de répondre aux besoins du quotidien.

À l’époque, j’utilisais l’éditeur de code d’Excel : chaque action dans le tableur était traduite automatiquement en code.  Je décortiquais ces transcriptions pour comprendre quelle instruction correspondait à chaque manipulation, puis je les combinais pour construire pas à pas un petit programme comptable.

De fil en aiguille, je découvrais les spécificités de VBA jusqu’à pouvoir écrire directement le code de manière autonome.

---

## Difficultés rencontrées

À cette époque, l’accès à l’information était bien plus limité qu’aujourd’hui. Les ressources disponibles en ligne étaient moins nombreuses, ce qui rendait l’apprentissage et le développement plus laborieux.

---

## Explications

Le fichier **Clients_.bas** contient plusieurs procédures (Sub) destinées à être associées à des **boutons personnalisés** dans la barre d’outils (fonction disponible uniquement sur la version PC, absente de Microsoft Office 2011).


### Liste des procédures

- `Sub Bouton_Nouveau_Dossier_Clients()`
- `Sub Bouton_Switch_Facture_Devis_Facture_et_Livraison_1_Page()`
- `Sub Bouton_Mise_à_Jour_Facture_et_Livraison_1_Page()`
- `Sub Bouton_finition_facture_acompte()`
- `Sub Bouton_Impression_Facture_et_Livraison_1_Page()`
- `Sub Bouton_Nouvelle_Facture_Livraison_Certificat_1_Page()`


### Ordre d’exécution conseillé

1. `Bouton_Nouveau_Dossier_Clients()`
   Crée un dossier client.

2. `Bouton_Mise_à_Jour_Facture_et_Livraison_1_Page()`
   Transfère les informations insérées dans la feuille Adresse vers les autres feuilles.

3. `Bouton_Switch_Facture_Devis_Facture_et_Livraison_1_Page()`
   Permet de transformer une facture en devis (et inversement) selon le besoin.

4. `Bouton_finition_facture_acompte()`
   Ajoute l’acompte versé et calcule le solde restant.

5. `Bouton_Impression_Facture_et_Livraison_1_Page()`
   IImprime la ou les pages de la facture et du bon de livraison.

6. `Nouvelle_Facture_Livraison_Certificat_1_Page()`
   Crée une nouvelle facture (de une à trois pages).


### Autres fichiers
Les fichiers **XLSX** fournis dans le repository sont des **aperçus visuels** des résultats générés par les macros contenues dans le fichier **Clients_.bas**. 

---

## Prochaine étape  

Continuer à perfectionner ce programme **Client** avec des mises à jour régulières, et amorcer le développement d’un programme de gestion des salaires.  





