# DEV-Rituels

Application Google Apps Script pour la gestion des rituels de l'association Les Jeux Démoniaques.

## Structure du Projet

```
DEV-Rituels/
├── appsscript.json    # Configuration du projet Apps Script
├── audit.js           # Système d'audit des actions
├── formula.js         # Fonctions de calcul
├── helpers.js         # Fonctions utilitaires
├── init.js           # Initialisation
├── logger.js         # Logging
├── main.js           # Point d'entrée principal
├── menu.js           # Menu de l'application
├── paypal.js         # Gestion des paiements
└── webapp.js         # Points d'entrée Web
```

## Configuration

1. Installer clasp :
   ```bash
   npm install -g @google/clasp
   ```

2. Se connecter à Google :
   ```bash
   clasp login
   ```

3. Pousser le code :
   ```bash
   clasp push
   ```

4. Configurer les propriétés du script (File > Project Settings > Script Properties) :
   - `ENV` : "dev" pour les tests, "prod" pour la production
   - `PAYPAL_CLIENT_ID` : Client ID PayPal
   - `PAYPAL_CLIENT_SECRET` : Secret PayPal
   - `PAYPAL_BRAND_NAME` : Nom de la marque/évènement
   - `PAYMENT_DEADLINE_TEXTE` : Texte pour la date limite de paiement
   - `EVENT_NAME` : Texte pour le nom de l'évenement dans le titre du mail
   - `EVENT_DATE_TEXT` : Texte pour la date de l'évènement
   - `EVENT_HOURS_TEXT` : Texte pour les horaires de l'évènement
   - `PAIEMENT_OPEN` : "true" si les paiements sont ouverts, "false" sinon
   - `WEBAPP_URL` : URL de l'application web déployée
   - `WRAPPER_URL` : URL du wrapper (interface utilisateur)

## Utilisation

L'application est déployée via Google Apps Script et utilise :
- Google Sheets pour les données
- PayPal pour les paiements
- Gmail pour l'envoi des mails