📌 Présentation du projet – Outil de création automatisée de groupes Microsoft 365 dynamiques
🎯 Objectif du projet

Ce projet a pour but de simplifier, sécuriser et standardiser la création de groupes Microsoft 365 dynamiques (Teams / M365 Groups) à partir de règles RH et organisationnelles, sans intervention manuelle dans Azure AD.

Il fournit :

une interface graphique (GUI) simple pour l’utilisateur,

une automatisation complète via Microsoft Graph,

une traçabilité SharePoint,

et une gestion d’erreurs claire et visible.

🧩 Architecture globale

Le projet est composé de 2 scripts principaux :

Script	Rôle
GUI.ps1	Interface graphique pour l’utilisateur
Create.ps1	Création technique du groupe dans Microsoft 365

Les deux scripts communiquent via des fichiers temporaires JSON / TXT.

🖥️ 1️⃣ GUI.ps1 – Interface utilisateur
🎨 Rôle du GUI

Le GUI est une application Windows Forms qui permet à un utilisateur non technique de :

Sélectionner les critères du groupe

Visualiser le nom du groupe avant création

Lancer la création sans connaître Azure AD ni Graph

Voir immédiatement le résultat (succès ou erreur)

🔗 Connexions utilisées

SharePoint Online (PnP PowerShell)
→ pour récupérer les dictionnaires métiers :

Produits

Sous-produits

Départements

Sous-départements (roles)

Pays

Villes

Ces listes garantissent que les valeurs utilisées sont officielles et normalisées.

🧠 Logique métier intégrée

L’utilisateur choisit :

Produit / Sous-produit

Département / Sous-département

Pays / Ville

Options :

Inclure ou non les externes

Managers uniquement ou non

À partir de ces choix, le script :

✅ Génère automatiquement :

DisplayName du groupe

MailNickname (conforme aux règles M365)

Règle de membership dynamique Azure AD

Exemple de règle :

(user.extensionAttribute1 -eq "Paris")
-and (user.extensionAttribute4 -eq "Finance")
-and (user.extensionAttribute10 -contains ",treatAsEmployee")

📄 Génération du fichier JSON

Avant création, le GUI génère un fichier :

C:\Temp\groupData.json


Contenant :

{
  "displayName": "...",
  "mailNickname": "...",
  "description": "...",
  "membershipRule": "..."
}


Ce fichier est ensuite transmis au script technique.

🚀 Lancement du script de création

Le GUI lance automatiquement :

Create.ps1


Puis :

Attend la fin de l’exécution

Lit le résultat

Affiche le succès ou l’erreur dans l’interface

🧾 Journalisation (Audit)

En cas de succès, le GUI écrit dans une liste SharePoint “log” :

Nom du groupe

Créateur

Date

Règle dynamique

Adresse mail du groupe

👉 Cela garantit une traçabilité complète.

⚙️ 2️⃣ Create.ps1 – Création technique du groupe
🔐 Connexion sécurisée

Connexion App Registration (client secret) à Microsoft Graph

Mode App-only (pas dépendant de l’utilisateur)

📥 Lecture des données

Lecture du fichier groupData.json

Vérification de son existence et de son contenu

Gestion d’erreurs immédiate si invalide

🧱 Création du groupe

Le script crée un :

Microsoft 365 Group (Unified)

Non sécurisé

Avec messagerie activée

🔄 Conversion en groupe dynamique

Une fois créé, le groupe est :

Converti en Dynamic Membership

La règle dynamique est activée

Le traitement automatique est mis sur ON

❗ Gestion des erreurs (améliorée)

Chaque étape est protégée par des try / catch :

Étape	Erreur capturée
Connexion Graph	Problème d’authentification
Lecture JSON	JSON manquant ou invalide
Création groupe	Groupe existant / droits
Conversion dynamique	Règle invalide / conflit

Le résultat est écrit dans :

C:\Temp\creationStatus.txt


Exemple :

CRE=0
ERROR=Erreur conversion dynamique : Groupe déjà existant

📤 Retour vers le GUI

Le script fournit :

Le statut (succès / échec)

Le message d’erreur détaillé

Le compte Graph utilisé

Le GUI affiche ces informations directement à l’utilisateur.

✅ Bénéfices du projet

✔ Standardisation des groupes
✔ Zéro erreur humaine sur les règles
✔ Accessible aux non-techniciens
✔ Audit et traçabilité complets
✔ Automatisation Graph sécurisée
✔ Interface claire avec feedback immédiat

🏁 Conclusion

Ce projet est un outil enterprise-ready qui transforme une opération complexe (création de groupes dynamiques Azure AD) en un processus simple, contrôlé et sécurisé, tout en respectant les règles métier et les standards IT.

👉 Il est parfaitement adapté à un environnement corporate Microsoft 365 à grande échelle.
