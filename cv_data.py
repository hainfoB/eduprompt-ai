# -*- coding: utf-8 -*-
"""
Full CV / credibility content for the public /about page.
Kept separate from translations.py (UI chrome) since this is narrative
resume content, already sourced in AR/FR/EN from the founder's documents.
"""

CV = {

    "fr": {
        "name": "Ahmed Haithem BERKANE",
        "title": "Architecte Logiciel & IA | Consultant Senior IT | Expert en Ingénierie Pédagogique (CIP)",
        "tagline": "Facilitateur de la transition numérique par l'Intelligence Artificielle et la Data.",
        "location": "Bordj Bou Arréridj, Algérie",
        "objective": "Consultant et développeur passionné par l'innovation, l'intelligence artificielle et l'éducation. Mon objectif : mettre mon expertise technique et pédagogique au service de projets à impact réel.",

        "stats": [
            {"value": "1000+", "label": "enseignants et cadres formés"},
            {"value": "56K+", "label": "abonnés communauté Instagram"},
            {"value": "700+", "label": "heures de formation en ligne"},
            {"value": "15", "label": "étudiants étrangers encadrés (Syrie, Irak)"},
        ],

        "section_experience": "Parcours professionnel",
        "experience": [
            {"period": "2021 — Présent", "title": "Consultant Indépendant & Expert IA", "org": "Freelance",
             "bullets": ["Conception d'architectures RAG et de modèles de langage privés pour entreprises",
                         "Plus de 1000 cadres et étudiants algériens formés, ainsi que 15 étudiants étrangers",
                         "Mentorat continu d'étudiants de la diaspora algérienne (France, Allemagne, Belgique, Canada)"]},
            {"period": "2021 — Présent", "title": "Formateur & Consultant", "org": "Freelance",
             "bullets": ["Mentorat de startups et de projets étudiants",
                         "Plus de 700 heures de formation en ligne (2020-2025), dont une formation gratuite pour 1000+ étudiants en 2025"]},
            {"period": "2021 — Présent", "title": "Instructeur vocationnel", "org": "CFPA BBA 04",
             "bullets": ["Enseignement de modules techniques aux techniciens en formation professionnelle",
                         "Conception de contenus et programmes pédagogiques"]},
            {"period": "2019 — 2020", "title": "Directeur de Projet Métier", "org": "Condor Academy",
             "bullets": ["Pilotage et obtention de l'agrément officiel du Ministère de l'Enseignement Supérieur",
                         "Direction de la mise en place de l'ERP SAP FICO et déploiement du LMS Success Factors",
                         "Conception et développement du site web officiel de l'académie et de la plateforme Microsoft Academy"]},
            {"period": "Mar. — Jul. 2020", "title": "Responsable par intérim des services de formation", "org": "CFPA BBA 04",
             "bullets": ["Mise en place d'une plateforme e-learning et d'une chaîne YouTube",
                         "Supervision des examens finaux pour écoles professionnelles privées"]},
            {"period": "Mar. — Jul. 2021", "title": "Responsable par intérim de la formation sur site", "org": "INSFP BBA 01",
             "bullets": ["Gestion des opérations pédagogiques et administratives",
                         "Coordination avec la direction des études"]},
            {"period": "2014 — 2020", "title": "Instructeur de laboratoire", "org": "Université de Bordj Bou Arréridj",
             "bullets": ["Cours : Génie logiciel, Delphi, Algorithmique, Systèmes d'information"]},
            {"period": "2014 — 2021", "title": "Professeur Spécialisé & CIP (PSFEP 2)", "org": "Secteur de la formation professionnelle",
             "bullets": ["Ingénierie pédagogique de haut niveau et conception de programmes de formation continue"]},
        ],

        "section_education": "Formation académique",
        "education": [
            {"period": "2013", "title": "Master en Informatique — Réseaux & Multimédia", "org": "Université El Bachir Ibrahimi, Bordj Bou Arréridj",
             "detail": "Poster de recherche : HTTP Digest dans les réseaux VoIP"},
            {"period": "2011", "title": "Licence en Informatique décisionnelle", "org": "Université El Bachir Ibrahimi, BBA",
             "detail": "Major de promotion ; projet classé 4ᵉ au niveau national (classification de patients par Machine Learning)"},
            {"period": "2006 — 2008", "title": "Classe préparatoire ingénieur", "org": "ENPEI Rouiba, Alger",
             "detail": "Formation militaire de base incluse"},
            {"period": "2006", "title": "Baccalauréat — Sciences de la nature", "org": "Lycée Abdelhamid Akrouf",
             "detail": "Moyenne 13.03 — 4ᵉ position"},
            {"period": "2015", "title": "Certificat de formation pédagogique", "org": "IFPG Guenbour, Sétif", "detail": ""},
        ],

        "section_skills": "Domaines d'expertise",
        "skills_groups": [
            {"title": "Intelligence Artificielle & Gouvernance",
             "skills": ["Prompt Engineering (institutionnel, industriel)", "IA Générative (LLM) & Architecture RAG",
                       "Machine Learning & Modélisation prédictive", "Data Analysis & Business Intelligence (Plotly)",
                       "Éthique de l'IA & Souveraineté numérique"]},
            {"title": "Ingénierie & Pédagogie",
             "skills": ["Architecture Full-Stack (Laravel, Python)", "Déploiement de systèmes ERP (SAP FICO)",
                       "Audit de sécurité & gouvernance des données", "Ingénierie pédagogique (certification CIP)",
                       "Pilotage de projets institutionnels complexes"]},
            {"title": "Développement technique",
             "skills": ["PHP, Laravel, MySQL, Bootstrap, HTML/CSS, JavaScript", "Delphi, Java, C#, Python",
                       "SQL Server, Oracle, Access", "Windows, Linux, LaTeX"]},
        ],

        "section_projects": "Réalisations logicielles",
        "projects": [
            {"title": "ASR PRO — Numérisation des examens", "desc": "Système complet de gestion et d'évaluation sécurisée de concours en ligne, présenté comme modèle de transformation digitale pour les instituts de formation."},
            {"title": "Moufid BI — Intelligence financière", "desc": "Plateforme de diagnostic financier pour groupes industriels : analyse automatisée de la performance et contrôle de gestion prédictif."},
            {"title": "Plateforme Zakat Fitr", "desc": "Conception et développement d'une plateforme numérique citoyenne dédiée à la gestion de la Zakat."},
            {"title": "Simulateur Python & C (2026)", "desc": "Environnement de simulation interactif conçu pour l'apprentissage de la programmation."},
        ],

        "section_influence": "Rayonnement & influence numérique",
        "influence": [
            "Créateur de contenu digital : communauté éducative de plus de 56 000 abonnés Instagram et 12 000 sur Facebook",
            "Création de cours interactifs en Darija (1ère initiative du genre) pour démocratiser l'accès aux technologies",
            "Préparation au Doctorat (PhD) : 25 sessions en ligne depuis 2021",
            "Data & Programmation : 5 sessions en ligne (Python, Analyse de données, IA), avec 15 étudiants étrangers (Syrie, Irak)",
            "Accompagnement de la diaspora algérienne (France, Allemagne, Belgique, Canada) dans leurs thèses et PFE",
        ],

        "section_events": "Engagements & conférences récents (2025-2026)",
        "events": [
            {"date": "Mars 2026", "title": "Global Africa Tech", "desc": "Participation au grand sommet de la tech et de l'innovation en Afrique — souveraineté numérique locale et continentale."},
            {"date": "Nov. 2025", "title": "Semaine Mondiale de l'Entrepreneuriat", "desc": "Conférencier sur l'intégration de l'intelligence artificielle dans les startups."},
            {"date": "Oct. 2025", "title": "Salon du Métier", "desc": "Exposition du projet ASR PRO et accompagnement des stagiaires INSFP BBA 01 sur l'intégration de l'IA (Business Model Canvas)."},
            {"date": "Sept. 2025", "title": "Forum National des Créateurs de Contenu Numérique", "desc": "Participation à la proposition d'un statut indépendant du créateur de contenu digital."},
        ],

        "section_certifications": "Certifications & reconnaissances",
        "certifications": [
            "Attestation de succès — Master en Informatique, Université de Bordj Bou Arréridj",
            "Certificat de reconnaissance — Membre du jury, Olympiades des Métiers (Skills Olympics) 2025",
            "Certificat de remerciement — Participation au Salon International de la Créativité et de l'Innovation, Centre Aïcha Haddad 2024",
            "Certificate of Speaking — WTM & GDG BBA, Journée Internationale de la Femme 2020",
            "Certificat de reconnaissance — Youth Elite of Sciences Association, 2020",
        ],

        "section_languages": "Langues",
        "languages": [
            {"lang": "Arabe", "level": "Langue maternelle"},
            {"lang": "Français", "level": "Niveau avancé"},
            {"lang": "Anglais", "level": "Niveau avancé"},
        ],
    },

    "en": {
        "name": "Ahmed Haithem BERKANE",
        "title": "Software & AI Architect | Senior IT Consultant | Pedagogical Engineering Expert (CIP)",
        "tagline": "Facilitating digital transformation through Artificial Intelligence and Data.",
        "location": "Bordj Bou Arréridj, Algeria",
        "objective": "Consultant and developer passionate about innovation, artificial intelligence, and education. My goal: put my technical and pedagogical expertise to work on projects with real impact.",

        "stats": [
            {"value": "1000+", "label": "teachers and professionals trained"},
            {"value": "56K+", "label": "Instagram community followers"},
            {"value": "700+", "label": "hours of online teaching"},
            {"value": "15", "label": "international students mentored (Syria, Iraq)"},
        ],

        "section_experience": "Professional Experience",
        "experience": [
            {"period": "2021 — Present", "title": "Independent Consultant & AI Expert", "org": "Freelance",
             "bullets": ["Designing RAG architectures and private language models for businesses",
                         "Over 1000 Algerian professionals and students trained, plus 15 international students",
                         "Ongoing mentorship of Algerian diaspora students (France, Germany, Belgium, Canada)"]},
            {"period": "2021 — Present", "title": "Trainer & Consultant", "org": "Freelance",
             "bullets": ["Mentoring startups and student projects",
                         "Over 700 hours of online teaching (2020-2025), including free training for 1000+ students in 2025"]},
            {"period": "2021 — Present", "title": "Vocational Instructor", "org": "CFPA BBA 04",
             "bullets": ["Teaching technical modules to vocational trainees",
                         "Designing pedagogical content and curricula"]},
            {"period": "2019 — 2020", "title": "Business Project Director", "org": "Condor Academy",
             "bullets": ["Led the process to obtain official accreditation from the Ministry of Higher Education",
                         "Directed the SAP FICO ERP rollout and the Success Factors LMS deployment",
                         "Designed and developed the academy's official website and the Microsoft Academy platform"]},
            {"period": "Mar. — Jul. 2020", "title": "Acting Head of Training Services", "org": "CFPA BBA 04",
             "bullets": ["Implemented an e-learning platform and a YouTube channel",
                         "Oversaw final exams for private vocational schools"]},
            {"period": "Mar. — Jul. 2021", "title": "Interim Head of On-Site Training", "org": "INSFP BBA 01",
             "bullets": ["Managed pedagogical and administrative operations",
                         "Coordinated with the deputy director of studies"]},
            {"period": "2014 — 2020", "title": "Lab Instructor", "org": "University of Bordj Bou Arréridj",
             "bullets": ["Courses: Software Engineering, Delphi, Algorithmics, Information Systems"]},
            {"period": "2014 — 2021", "title": "Specialized Teacher & CIP (PSFEP 2)", "org": "Vocational training sector",
             "bullets": ["High-level pedagogical engineering and continuing-education program design"]},
        ],

        "section_education": "Education",
        "education": [
            {"period": "2013", "title": "Master's in Computer Science — Networks & Multimedia", "org": "University El Bachir Ibrahimi, Bordj Bou Arréridj",
             "detail": "Research poster: HTTP Digest in VoIP Networks"},
            {"period": "2011", "title": "Bachelor's in Decision-Support Computer Science", "org": "University El Bachir Ibrahimi, BBA",
             "detail": "Top of class; project ranked 4th nationally (patient classification via Machine Learning)"},
            {"period": "2006 — 2008", "title": "Preparatory Engineering Program", "org": "ENPEI Rouiba, Algiers",
             "detail": "Included basic military training"},
            {"period": "2006", "title": "Baccalaureate — Natural Sciences", "org": "Abdelhamid Akrouf High School",
             "detail": "Average 13.03 — ranked 4th"},
            {"period": "2015", "title": "Pedagogical Training Certificate", "org": "IFPG Guenbour, Sétif", "detail": ""},
        ],

        "section_skills": "Areas of Expertise",
        "skills_groups": [
            {"title": "Artificial Intelligence & Governance",
             "skills": ["Prompt Engineering (institutional, industrial)", "Generative AI (LLM) & RAG Architecture",
                       "Machine Learning & Predictive Modeling", "Data Analysis & Business Intelligence (Plotly)",
                       "AI Ethics & Digital Sovereignty"]},
            {"title": "Engineering & Pedagogy",
             "skills": ["Full-Stack Architecture (Laravel, Python)", "ERP System Deployment (SAP FICO)",
                       "Security Audit & Data Governance", "Pedagogical Engineering (CIP Certification)",
                       "Complex Institutional Project Management"]},
            {"title": "Technical Development",
             "skills": ["PHP, Laravel, MySQL, Bootstrap, HTML/CSS, JavaScript", "Delphi, Java, C#, Python",
                       "SQL Server, Oracle, Access", "Windows, Linux, LaTeX"]},
        ],

        "section_projects": "Software Achievements",
        "projects": [
            {"title": "ASR PRO — Exam Digitization", "desc": "Complete system for secure online exam management and assessment, presented as a digital transformation model for training institutes."},
            {"title": "Moufid BI — Financial Intelligence", "desc": "Financial diagnostic platform for industrial groups: automated performance analysis and predictive management control."},
            {"title": "Zakat Fitr Platform", "desc": "Design and development of a citizen-facing digital platform dedicated to managing Zakat."},
            {"title": "Python & C Simulator (2026)", "desc": "Interactive simulation environment designed for programming education."},
        ],

        "section_influence": "Digital Reach & Influence",
        "influence": [
            "Digital content creator: educational community of over 56,000 Instagram followers and 12,000 on Facebook",
            "Created interactive courses in Darija (a first of its kind) to democratize access to technology",
            "PhD preparation: 25 online sessions since 2021",
            "Data & Programming: 5 online sessions (Python, Data Analysis, AI), with 15 international students (Syria, Iraq)",
            "Supporting the Algerian diaspora (France, Germany, Belgium, Canada) with their theses and capstone projects",
        ],

        "section_events": "Recent Engagements & Conferences (2025-2026)",
        "events": [
            {"date": "Mar. 2026", "title": "Global Africa Tech", "desc": "Participation in Africa's major tech and innovation summit — focus on local and continental digital sovereignty."},
            {"date": "Nov. 2025", "title": "Global Entrepreneurship Week", "desc": "Speaker on integrating artificial intelligence into startups."},
            {"date": "Oct. 2025", "title": "Salon du Métier (Trade Fair)", "desc": "Showcased the ASR PRO project and mentored INSFP BBA 01 trainees on AI integration (Business Model Canvas)."},
            {"date": "Sep. 2025", "title": "National Forum of Digital Content Creators", "desc": "Contributed to the proposal for an independent status for digital content creators."},
        ],

        "section_certifications": "Certifications & Recognitions",
        "certifications": [
            "Certificate of Success — Master's in Computer Science, University of Bordj Bou Arréridj",
            "Recognition Certificate — Jury Member, Skills Olympics 2025",
            "Certificate of Appreciation — International Creativity & Innovation Fair, Aïcha Haddad Center, 2024",
            "Certificate of Speaking — WTM & GDG BBA, International Women's Day 2020",
            "Recognition Certificate — Youth Elite of Sciences Association, 2020",
        ],

        "section_languages": "Languages",
        "languages": [
            {"lang": "Arabic", "level": "Native"},
            {"lang": "French", "level": "Advanced"},
            {"lang": "English", "level": "Advanced"},
        ],
    },

    "ar": {
        "name": "أحمد هيثم بركان",
        "title": "مهندس برمجيات وذكاء اصطناعي | مستشار تقني أول | خبير هندسة بيداغوجية (CIP)",
        "tagline": "ميسّر للتحول الرقمي عبر الذكاء الاصطناعي والبيانات.",
        "location": "برج بوعريريج، الجزائر",
        "objective": "مستشار ومطوّر شغوف بالابتكار والذكاء الاصطناعي والتعليم. هدفي: تسخير خبرتي التقنية والبيداغوجية لخدمة مشاريع ذات أثر حقيقي.",

        "stats": [
            {"value": "+1000", "label": "أستاذ وإطار تم تكوينهم"},
            {"value": "+56 ألف", "label": "متابع على المجتمع التعليمي بإنستغرام"},
            {"value": "+700", "label": "ساعة تكوين عبر الإنترنت"},
            {"value": "15", "label": "طالباً أجنبياً تم تأطيرهم (سوريا، العراق)"},
        ],

        "section_experience": "المسار المهني",
        "experience": [
            {"period": "2021 — الحاضر", "title": "مستشار مستقل وخبير ذكاء اصطناعي", "org": "عمل حر",
             "bullets": ["تصميم معماريات RAG ونماذج لغوية خاصة للمؤسسات",
                         "أكثر من 1000 إطار وطالب جزائري تم تكوينهم، بالإضافة إلى 15 طالباً أجنبياً",
                         "تأطير مستمر لطلبة الجالية الجزائرية (فرنسا، ألمانيا، بلجيكا، كندا)"]},
            {"period": "2021 — الحاضر", "title": "مكوّن ومستشار", "org": "عمل حر",
             "bullets": ["تأطير المشاريع الناشئة ومشاريع الطلبة",
                         "أكثر من 700 ساعة تكوين عبر الإنترنت (2020-2025)، بما فيها تكوين مجاني لأكثر من 1000 طالب في 2025"]},
            {"period": "2021 — الحاضر", "title": "مكوّن مهني", "org": "CFPA برج بوعريريج 04",
             "bullets": ["تدريس وحدات تقنية للمتربصين في التكوين المهني",
                         "تصميم محتويات وبرامج بيداغوجية"]},
            {"period": "2019 — 2020", "title": "مدير مشروع أعمال", "org": "أكاديمية كوندور",
             "bullets": ["قيادة والحصول على الاعتماد الرسمي من وزارة التعليم العالي والبحث العلمي",
                         "قيادة تنصيب نظام ERP SAP FICO ونشر منصة LMS Success Factors",
                         "تصميم وتطوير الموقع الرسمي للأكاديمية ومنصة Microsoft Academy"]},
            {"period": "مارس — جويلية 2020", "title": "مسؤول بالنيابة لخدمات التكوين", "org": "CFPA برج بوعريريج 04",
             "bullets": ["تنصيب منصة تعليم إلكتروني وقناة يوتيوب",
                         "الإشراف على الامتحانات النهائية للمدارس المهنية الخاصة"]},
            {"period": "مارس — جويلية 2021", "title": "مسؤول بالنيابة للتكوين الحضوري", "org": "INSFP برج بوعريريج 01",
             "bullets": ["تسيير العمليات البيداغوجية والإدارية",
                         "التنسيق مع نيابة مديرية الدراسات"]},
            {"period": "2014 — 2020", "title": "مكوّن مخبر", "org": "جامعة برج بوعريريج",
             "bullets": ["مقاييس: هندسة البرمجيات، Delphi، الخوارزميات، نظم المعلومات"]},
            {"period": "2014 — 2021", "title": "أستاذ متخصص ومستشار تكوين وتفتيش بيداغوجي (PSFEP 2)", "org": "قطاع التكوين المهني",
             "bullets": ["هندسة بيداغوجية رفيعة المستوى وتصميم برامج التكوين المستمر"]},
        ],

        "section_education": "التكوين الأكاديمي",
        "education": [
            {"period": "2013", "title": "ماستر في الإعلام الآلي — شبكات ووسائط متعددة", "org": "جامعة الأمير عبد القادر / الباشير الإبراهيمي، برج بوعريريج",
             "detail": "ملصق بحثي: HTTP Digest في شبكات VoIP"},
            {"period": "2011", "title": "ليسانس في الإعلام الآلي القرار", "org": "جامعة الباشير الإبراهيمي، برج بوعريريج",
             "detail": "الأول على الدفعة؛ مشروع مصنف 4 وطنياً (تصنيف المرضى بالتعلم الآلي)"},
            {"period": "2006 — 2008", "title": "تحضيري هندسة", "org": "ENPEI الرويبة، الجزائر العاصمة",
             "detail": "شمل تكويناً عسكرياً أساسياً"},
            {"period": "2006", "title": "بكالوريا — علوم طبيعية", "org": "ثانوية عبد الحميد عكروف",
             "detail": "معدل 13.03 — الرتبة 4"},
            {"period": "2015", "title": "شهادة تكوين بيداغوجي", "org": "IFPG قنبور، سطيف", "detail": ""},
        ],

        "section_skills": "مجالات الخبرة",
        "skills_groups": [
            {"title": "الذكاء الاصطناعي والحوكمة",
             "skills": ["هندسة الأوامر (مؤسساتي، صناعي)", "الذكاء الاصطناعي التوليدي (LLM) ومعمارية RAG",
                       "التعلم الآلي والنمذجة التنبؤية", "تحليل البيانات وذكاء الأعمال (Plotly)",
                       "أخلاقيات الذكاء الاصطناعي والسيادة الرقمية"]},
            {"title": "الهندسة والبيداغوجيا",
             "skills": ["معمارية Full-Stack (Laravel، Python)", "نشر أنظمة ERP (SAP FICO)",
                       "تدقيق الأمن وحوكمة البيانات", "الهندسة البيداغوجية (شهادة CIP)",
                       "قيادة المشاريع المؤسساتية المعقدة"]},
            {"title": "التطوير التقني",
             "skills": ["PHP، Laravel، MySQL، Bootstrap، HTML/CSS، JavaScript", "Delphi، Java، C#، Python",
                       "SQL Server، Oracle، Access", "Windows، Linux، LaTeX"]},
        ],

        "section_projects": "الإنجازات البرمجية",
        "projects": [
            {"title": "ASR PRO — رقمنة الامتحانات", "desc": "نظام كامل لإدارة وتقييم المسابقات عبر الإنترنت بأمان، يُقدَّم كنموذج للتحول الرقمي لمعاهد التكوين."},
            {"title": "Moufid BI — الذكاء المالي", "desc": "منصة تشخيص مالي للمجموعات الصناعية: تحليل آلي للأداء ومراقبة تسيير تنبؤية."},
            {"title": "منصة زكاة الفطر", "desc": "تصميم وتطوير منصة رقمية مواطنة مخصصة لتسيير الزكاة."},
            {"title": "محاكي Python و C (2026)", "desc": "بيئة محاكاة تفاعلية مصممة لتعلم البرمجة."},
        ],

        "section_influence": "الإشعاع والتأثير الرقمي",
        "influence": [
            "صانع محتوى رقمي: مجتمع تعليمي يضم أكثر من 56 ألف متابع على إنستغرام و12 ألف على فيسبوك",
            "إنشاء دورات تفاعلية بالدارجة (مبادرة أولى من نوعها) لتعميم الوصول إلى التكنولوجيا",
            "التحضير للدكتوراه: 25 جلسة عبر الإنترنت منذ 2021",
            "البيانات والبرمجة: 5 جلسات عبر الإنترنت (Python، تحليل البيانات، الذكاء الاصطناعي)، مع 15 طالباً أجنبياً (سوريا، العراق)",
            "مرافقة الجالية الجزائرية (فرنسا، ألمانيا، بلجيكا، كندا) في أطروحاتهم ومشاريع تخرجهم",
        ],

        "section_events": "المشاركات والمؤتمرات الأخيرة (2025-2026)",
        "events": [
            {"date": "مارس 2026", "title": "Global Africa Tech", "desc": "المشاركة في القمة الكبرى للتكنولوجيا والابتكار في إفريقيا — السيادة الرقمية المحلية والقارية."},
            {"date": "نوفمبر 2025", "title": "الأسبوع العالمي لريادة الأعمال", "desc": "متحدث حول دمج الذكاء الاصطناعي في المؤسسات الناشئة."},
            {"date": "أكتوبر 2025", "title": "صالون المهنة", "desc": "عرض مشروع ASR PRO ومرافقة متربصي INSFP برج بوعريريج 01 حول دمج الذكاء الاصطناعي (Business Model Canvas)."},
            {"date": "سبتمبر 2025", "title": "الملتقى الوطني لصنّاع المحتوى الرقمي", "desc": "المساهمة في اقتراح وضع مستقل لصانع المحتوى الرقمي."},
        ],

        "section_certifications": "الشهادات والتكريمات",
        "certifications": [
            "شهادة نجاح — ماستر في الإعلام الآلي، جامعة برج بوعريريج",
            "شهادة تقدير — عضو لجنة تحكيم، أولمبياد المهن (Skills Olympics) 2025",
            "شهادة شكر — المشاركة في الصالون الدولي للإبداع والابتكار، مركز عائشة حداد 2024",
            "Certificate of Speaking — WTM و GDG برج بوعريريج، اليوم العالمي للمرأة 2020",
            "شهادة تقدير — جمعية نخبة الشباب للعلوم، 2020",
        ],

        "section_languages": "اللغات",
        "languages": [
            {"lang": "العربية", "level": "اللغة الأم"},
            {"lang": "الفرنسية", "level": "مستوى متقدم"},
            {"lang": "الإنجليزية", "level": "مستوى متقدم"},
        ],
    },
}
