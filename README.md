# EduPrompt AI — Plateforme SaaS Pédagogique

Plateforme complète avec authentification, abonnements et génération de documents
pédagogiques par IA (Gemini), personnalisables (couleurs, en-tête, sources).

---

## ✨ Fonctionnalités

- 🔐 **Authentification** : inscription / connexion sécurisée (mots de passe hashés)
- 📊 **Dashboard** : suivi du forfait, quota, historique des documents
- 💳 **Abonnements** : essai gratuit (14 jours / 5 documents), forfaits Pro/Premium
  activables par code de licence (vente manuelle : virement, WhatsApp, etc.)
- 🎨 **Design personnalisé** : jusqu'à 3 couleurs + couleur de texte, appliquées
  à l'en-tête et aux titres du document généré
- 🧑‍🏫 **En-tête enseignant** : tableau avec Nom & Prénom, Niveau, Palier, Matière
- 📎 **Sources façon NotebookLM** : ajout de liens web et upload de fichiers
  (PDF, Word, TXT) utilisés comme contexte pour l'IA
- 🌍 Export **Word (.docx)** et **PDF**
- 📱 Responsive (Bootstrap 5)
- 👑 **Panel admin** : génération de codes de licence

---

## ⚙️ Installation

```bash
pip install -r requirements.txt
python app.py
```

Accéder à : **http://localhost:5000**

La base de données SQLite (`eduprompt.db`) se crée automatiquement au premier lancement.

---

## 👑 Créer un compte administrateur

Après le premier lancement, exécutez :

```bash
python3 -c "
from app import app
from models import db, User, Subscription
from datetime import datetime, timedelta

with app.app_context():
    u = User(first_name='Admin', last_name='EduPrompt', email='admin@eduprompt.dz', role='admin')
    u.set_password('changeme123')
    db.session.add(u)
    db.session.flush()
    sub = Subscription(user_id=u.id, plan='premium', status='active',
                        expires_at=datetime.utcnow()+timedelta(days=3650), docs_limit=999999)
    db.session.add(sub)
    db.session.commit()
    print('Admin créé:', u.email)
"
```

Connectez-vous avec `admin@eduprompt.dz` / `changeme123`, puis accédez à
**/admin/licenses** pour générer des codes d'activation Pro/Premium à vendre
à vos utilisateurs.

⚠️ **Changez ce mot de passe immédiatement après la première connexion.**

---

## 💰 Modèle d'abonnement

Le système inclut :
- **Essai gratuit** : 5 documents / 14 jours à l'inscription
- **Pro / Premium** : illimité, activé par code de licence

**Aucune passerelle de paiement automatique** (Stripe, PayPal…) n'est connectée —
le modèle est conçu pour une vente manuelle (virement bancaire, CCP, WhatsApp) :
vous générez un code via `/admin/licenses` et l'envoyez au client après paiement.

Pour automatiser les paiements en ligne, il faudra intégrer une passerelle
(Stripe Checkout, CIB/Edahabia via un PSP local, etc.) — cela nécessite vos
propres clés API marchand.

---

## 🔑 Clé API Gemini

Chaque enseignant saisit sa propre clé dans l'interface (jamais stockée côté
serveur — l'appel à Gemini se fait directement depuis son navigateur).

Clé gratuite : https://aistudio.google.com/app/apikey

---

## 🌐 Déploiement en production

```bash
pip install gunicorn
gunicorn -w 4 -b 0.0.0.0:8000 app:app
```

Pensez à :
- Définir `SECRET_KEY` en variable d'environnement
- Utiliser PostgreSQL plutôt que SQLite en production (changer `SQLALCHEMY_DATABASE_URI`)
- Servir derrière HTTPS (Nginx + Let's Encrypt)

---

## 📁 Structure

```
eduprompt/
├── app.py                    ← Backend Flask (routes, génération docx/pdf)
├── models.py                 ← Base de données (User, Subscription, Document, LicenseCode)
├── utils.py                  ← Extraction de sources (fichiers + liens)
├── requirements.txt
└── templates/
    ├── base.html              ← Layout partagé + navigation
    ├── login.html / register.html
    ├── dashboard.html         ← Suivi abonnement + historique
    ├── upgrade.html           ← Activation code de licence
    ├── admin_licenses.html    ← Génération de codes (admin)
    └── generator.html         ← Outil principal (couleurs, sources, génération)
```
