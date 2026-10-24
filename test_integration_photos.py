"""
Test d'intégration complet pour la gestion des photos de joueurs et les podiums RCA.
"""
import os
import sys
import pandas as pd
import photo_manager
import app

def test_photos():
    print("--- 1. Test des correspondances de photos ---")
    test_cases = [
        ("Landry SALUA", "landry_salua.png"),
        ("SALUA Landry", "landry_salua.png"),
        ("Clément BAMIERE", "clement_bamiere.png"),
        ("Bam's", "clement_bamiere.png"),
        ("BAM'S", "clement_bamiere.png"),
        ("Jean-Étienne LESUEUR", "jean_etienne_lesueur.png"),
        ("Jean-Étienne ( TIJ)", "jean_etienne_lesueur.png"),
        ("Maka SALUA", "maka_salua.png"),
        ("Maka", "maka_salua.png"),
        ("Arthur MOINE", "arthur_moine.png"),
        ("Enzo VASELLI", "enzo_vaselli.png"),
        ("Thomas WAEYAERT", "thomas_waeyaert.png"),
    ]
    for name, expected_file in test_cases:
        path = app.obtenir_chemin_photo(name)
        assert path is not None, f"Photo introuvable pour {name}"
        assert os.path.basename(path).lower() == expected_file.lower(), f"Mauvais fichier pour {name}: {path} vs {expected_file}"
        print(f"  [OK] {name} -> {os.path.basename(path)}")

    print("\n--- 2. Test get_player_image_b64 ---")
    b64_salua = app.get_player_image_b64("Landry SALUA", size=80)
    assert b64_salua.startswith("data:image/png;base64,"), "Format Base64 invalide pour joueur avec photo"
    print(f"  [OK] Landry SALUA Base64 len: {len(b64_salua)}")

    b64_inconnu = app.get_player_image_b64("Inconnu Test", size=80)
    assert b64_inconnu.startswith("data:image/svg+xml;base64,"), "Format SVG avatar invalide pour joueur sans photo"
    print(f"  [OK] Joueur inconnu SVG Base64 len: {len(b64_inconnu)}")

    print("\n--- 3. Test generer_podium_dynamique et HTML ---")
    with open("ENTRAINEMENT 080926.xlsx", "rb") as f_db:
        df_raw = app.load_catapult_file(f_db)
    df_indiv = app.process_catapult_data(df_raw, split_mode="session")
    podium = app.generer_podium_dynamique(df_indiv)

    assert "Top speed (km/h)" in podium
    assert "% de marche" in podium
    assert "Haute intensité (m)" in podium

    for cat_name, cat_data in podium.items():
        players = cat_data["players"]
        assert len(players) == 3
        print(f"\n  Podium: {cat_name}")
        for p in players:
            print(f"    #{p['rank']} {p['name']} ({p.get('poste')}) : {p['stat']}")
            assert "avatar_src" in p
            assert p["avatar_src"].startswith("data:image/")

    # Test markup HTML
    for cat_name, cat_data in podium.items():
        html = photo_manager.render_podium_html_markup(
            category_title=cat_name,
            category_icon=cat_data["icon"],
            players=cat_data["players"],
        )
        assert "border: 3.5px solid #FFD700" in html  # 1er Or
        assert "border: 3px solid #C0C0C0" in html    # 2e Argent
        assert "border: 3px solid #CD7F32" in html    # 3e Bronze
        assert "👑" in html
        assert cat_name in html

    print("\n  [OK] HTML Podium validé avec succès (marches Or, Argent, Bronze).")

    print("\n--- 4. Test Génération PDF 7 Pages avec Photos ---")
    pdf_bytes = app.generer_rapport_pdf_coachs(
        df_indiv=df_indiv,
        date_seance="26/10/2020",
    )
    assert len(pdf_bytes) > 50000, "Le PDF généré est anormalement petit"
    print(f"  [OK] PDF généré avec succès, taille: {len(pdf_bytes):,d} octets.")

    print("\n--- 5. Test Sauvegarde et Suppression Propre de Photo ---")
    dummy_nom = "TEST_JOUEUR_PHOTO"
    dummy_bytes = b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x06\x00\x00\x00\x1f\x15c4\x00\x00\x00\nIDATx\x9cc\x00\x01\x00\x00\x05\x00\x01\r\n-\xb4\x00\x00\x00\x00IEND\xaeB`\x82"
    ok_save, _ = app.sauvegarder_joueur_personnalise(
        nom=dummy_nom,
        poste="Pilier",
        groupe="5 de devant",
        photo_bytes=dummy_bytes,
        photo_ext=".png"
    )
    assert ok_save
    saved_path = app.obtenir_chemin_photo(dummy_nom)
    assert saved_path is not None and os.path.isfile(saved_path), "Fichier photo non trouvé après sauvegarde"
    print(f"  [OK] Photo sauvegardée: {saved_path}")

    ok_del, _ = app.supprimer_joueur_personnalise(dummy_nom)
    assert ok_del
    assert not os.path.exists(saved_path), "La photo orpheline n'a pas été supprimée !"
    print(f"  [OK] Photo supprimée proprement sans orphelin.")

    print("\n==========================================")
    print(" TOUS LES TESTS SONT PASSÉS AVEC SUCCÈS ! ")
    print("==========================================")

if __name__ == "__main__":
    test_photos()
