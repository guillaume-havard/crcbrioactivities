CRC Brio Activities
===================

Crunch Brio and crc activities files to generate errors and pre-fac files.

Analyse
-------
Analyse des données du fichier activité CRC

1. Vérifier que les palettes reçues à 0 soient notées en statut « annulée ».
Vérifier que toutes les commandes en statut « annulée » soient en palettes reçues à 0.
 
2. Sélectionner les lignes sans « tDate livraison réelle » pour vérifier le type d’incident dont il s’agit.
Si la date n’est pas présente ou inférieur au 1 janvier 2000 et que la livraison est en « livrée »,
mettre une alerte.

3. Associer les BL (_2 et _RET) au BL original pour vérifier le traitement de l’incident.

