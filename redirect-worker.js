/**
 * Retires resume.bastion-fold-site.workers.dev.
 *
 * The resume is published on its own site, deliberately separate from the
 * company. This company-branded host is neither separate nor nested (it has
 * no back-navigation to Bastion Fold), so it only forwards.
 *
 * Path and query are preserved: the Worker and Pages served the identical
 * _site build, so every deep link (including /downloads/*) still resolves.
 */

const CANONICAL_HOST = "alexander-sorrell-resume.pages.dev";

export default {
  fetch(request) {
    const url = new URL(request.url);
    url.protocol = "https:";
    url.hostname = CANONICAL_HOST;
    url.port = "";
    return Response.redirect(url.toString(), 301);
  },
};
