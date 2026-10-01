# Smooth scroll — obligatoire sur ce site

**Convention Pascal : tout site qu'il fait/qu'on fait pour lui DOIT avoir le scroll doux activé.** Ne pas le retirer, ne pas l'oublier sur une nouvelle page.

## Le fichier

`site/smooth-scroll.js` est copié à la création du projet par `newproj` depuis `/home/user/templates/smooth-scroll-kit/`. C'est un bundle all-in-one (Lenis + init + CSS injectée) — pas de build, pas de dépendance, ~20 ko.

## Comment l'activer (mode `data-auto`, à utiliser par défaut)

À mettre dans **chaque** page HTML du site, juste avant `</body>` :

```html
<script src="smooth-scroll.js" data-auto></script>
```

C'est tout. Le scroll devient doux dès le chargement.

## Si tu as besoin de tuner

```html
<script src="smooth-scroll.js"></script>
<script>
  SmoothScroll.init({
    lerp: 0.09,            // 0.05 = très doux / 0.09 = équilibré (par défaut) / 0.2 = plus réactif
    smoothWheel: true,
    wheelMultiplier: 1,
    touchMultiplier: 1
  });
</script>
```

`window.__lenis` est exposé ensuite si tu veux piloter (`__lenis.scrollTo('#section')`, `__lenis.stop()`, `__lenis.start()`).

## Accessibilité

Le kit **désactive automatiquement** le smooth-scroll si l'utilisateur a `prefers-reduced-motion: reduce`. Pas besoin de gérer ce cas à la main.

## Si tu utilises un framework / un sous-dossier `dist/`

- Build statique (Vite, Astro, …) : copie `smooth-scroll.js` dans le dossier publié (`public/`, `static/`, etc.) et garde la même `<script>` dans le `<body>` du template.
- React/Vue/Svelte SPA : importe ou charge le script une seule fois au boot ; ne pas l'inclure dans chaque composant.

## Si tu casses cette règle

Pascal va te le faire refaire. Ne déploie pas/ne livre pas une page sans ce script en place.

## Pour les chemins, animations dérivées du scroll (reveal, parallax, pinned)

Ils héritent automatiquement de l'inertie de Lenis — tu peux faire des `getBoundingClientRect()` + `IntersectionObserver` comme d'habitude, ça suivra le scroll doux. Voir le README du kit à `/home/user/templates/smooth-scroll-kit/README.md` pour des patterns (reveal-on-enter, scroll progress d'une section pinned, play-once forward-only).
