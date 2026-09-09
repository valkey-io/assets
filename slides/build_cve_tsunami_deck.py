#!/usr/bin/env python3
"""Build "Open Source Maintainers vs. The CVE Tsunami" from the placeholder deck.

    python3 slides/build_cve_tsunami_deck.py

Talk: Madelyn Olson (AWS) and Harkrishn Patro (Apple), Thursday 8 October 2026,
15:50-16:30 CEST, Club H (Floor 1). 40 minute slot, so ~30 minutes of content.

Content comes from the blog post "Keeping up with AI: Valkey security in 2026"
(valkey-io/valkey-io.github.io#664). Every number and every claim on a slide is from
that post; nothing here is invented. Sources are cited on the slides that carry data.

The donor is the placeholder deck, not the community deck: it is the one whose master
carries the eleven Valkey layouts. This script never touches the donor's master, so the
layouts and the patched theme come through untouched.
"""
import os
import shutil
import sys

from pptx import Presentation

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_placeholder_deck as V
from build_placeholder_deck import (LAYOUT_BIG_NUMBER, LAYOUT_INSIGHT,
                                    LAYOUT_INSIGHT_BOLD, LAYOUT_LIST_ROWS,
                                    LAYOUT_THANKS, LAYOUT_TITLE_BOTTOM,
                                    slide_bullets_cards, slide_from_layout,
                                    slide_three_box, slide_two_box)
from pptx.enum.dml import MSO_THEME_COLOR

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "Valkey placeholder slides.pptx")
OUT = os.path.join(HERE, "Open Source Maintainers vs The CVE Tsunami.pptx")

BLOG = "Source: valkey.io/blog/keeping-up-with-ai-valkey-security/"
CARD_A, CARD_B = MSO_THEME_COLOR.ACCENT_5, V.CARD_ROSE
NEUTRAL = MSO_THEME_COLOR.LIGHT_2


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text
    return slide


def build():
    shutil.copyfile(TEMPLATE, OUT + ".building")
    prs = Presentation(OUT + ".building")

    missing = [n for n in (LAYOUT_TITLE_BOTTOM, LAYOUT_INSIGHT, LAYOUT_INSIGHT_BOLD,
                           LAYOUT_BIG_NUMBER, LAYOUT_LIST_ROWS, LAYOUT_THANKS,
                           V.LAYOUT_TITLE_ONLY)
               if prs.slide_layouts.get_by_name(n) is None]
    if missing:
        raise SystemExit(f"{TEMPLATE} is missing layouts: {missing}")

    for i in range(len(prs.slides) - 1, -1, -1):          # highest index first
        prs.part.drop_rel(prs.slides._sldIdLst[i].rId)
        del prs.slides._sldIdLst[i]

    # --- 1. title ---------------------------------------------------------------
    notes(slide_from_layout(prs, LAYOUT_TITLE_BOTTOM, [
        "Open Source Maintainers vs. The CVE Tsunami",
        "How Valkey adapted when bugs got cheap to find",
        "Madelyn Olson, AWS  ·  Harkrishn Patro, Apple",
        "Club H (Floor 1)  ·  Thursday 8 October 2026, 15:50 CEST",
    ]), "Madelyn opens. Split: Madelyn takes the tsunami and the policy change, "
        "Harkrishn takes the adversarial audits and the release automation.")

    # --- 2. the rate changed, not the code --------------------------------------
    notes(slide_from_layout(prs, LAYOUT_BIG_NUMBER, [
        "7",
        "advisories in the first eight months of 2026",
        "Five in the previous 21 months, from the March 2024 fork. "
        "Plus almost three dozen security-adjacent fixes.",
    ]), "The point is the rate, not the total. Same project, same review standards, "
        "same maintainers. What changed is outside the project: the cost of finding "
        "a bug went to almost zero.")

    # --- 3. the duplicate-report tell -------------------------------------------
    # Deliberately not a second Big Number slide: this is a story, and two of those
    # back to back reads as a stat dump rather than two separate points.
    notes(slide_from_layout(prs, LAYOUT_INSIGHT, [
        "Three researchers, two weeks, one use-after-free in the script debugger.",
        "None of them knew about the others. That is the tell.",
    ]), "This is the tell that the reports are model-generated rather than the result "
        "of independent human audit. The script debugger is a feature almost nobody "
        "uses, so three people landing on it in a fortnight is not coincidence: it is "
        "what happens when everyone points a similar model at the same source tree.")

    # --- 4. not just us ---------------------------------------------------------
    notes(slide_from_layout(prs, LAYOUT_INSIGHT, [
        "A seemingly unending deluge of reports from researchers using LLMs to mine "
        "for security gold.",
        "Jeremy Stanley, OpenStack vulnerability management. Kernel, Red Hat and "
        "HAProxy maintainers report the same duplicates.",
    ]), "Read the quote, then say the names. The audience needs to know this is not a "
        "Valkey problem or a Valkey complaint. Links: openwall.com/lists/oss-security/"
        "2026/04/28/15 and lwn.net/Articles/1070698/")

    # --- 5. what one report actually costs --------------------------------------
    notes(slide_from_layout(prs, LAYOUT_LIST_ROWS, [
        "One report, five jobs, a handful of people",
        "Reproduce the report against the build it targets",
        "Judge the severity, and defend that judgement",
        "Write the fix",
        "Coordinate the embargo with downstream consumers",
        "Ship it across every supported version",
    ]), "None of these five scale with the number of reports. Four of them are "
        "maintainer judgement or coordination, which is exactly the part a model "
        "cannot hand off. This slide is why volume is the problem rather than "
        "difficulty.")

    # --- 6. the turn ------------------------------------------------------------
    notes(slide_from_layout(prs, LAYOUT_INSIGHT_BOLD, [
        "The industry celebrates AI finding more bugs. Maintainers are drowning in "
        "the reports.",
        "Both halves are true, and that is the whole problem.",
    ]), "The turn of the talk. Pause here. Everything before this is the problem, "
        "everything after is what we did. Do not soften it and do not rush off it.")

    # --- 7. the three changes ---------------------------------------------------
    notes(slide_three_box(
        prs, "Three changes that let us keep up",
        [("security", "Change what counts",
          "Stop spending coordination time on bugs that give an attacker nothing "
          "they did not already have."),
         ("hardening", "Find them first",
          "Run adversarial audits against our own code instead of waiting for the "
          "report to arrive."),
         ("migration", "Ship the fix faster",
          "Automate the backport and codify the release process, so a zero-day is a "
          "same-day fix.")],
        NEUTRAL,
    ), "Roadmap slide. Say the three, then say that none of them is about reading "
       "reports faster. Each one removes work rather than absorbing it.")

    # --- 8. what earns an advisory ----------------------------------------------
    notes(slide_bullets_cards(
        prs, "What earns an advisory now",
        [("Pre-auth", " crash from a malformed request. CVE-2026-27623 is this "
                      "year's example."),
         ("Out-of-bounds read", " reachable by an authenticated client, because the "
                                "bytes may belong to another client."),
         ("Memory corruption", " with a credible path to code execution."),
         "Any action beyond the permissions the client was granted"],
        source=BLOG,
    ), "The first three are the ones that give an attacker something new. The fourth "
       "is the general rule the other three are instances of: an advisory is for "
       "crossing a boundary, not for breaking something.")

    # --- 9. and what does not ---------------------------------------------------
    notes(slide_two_box(
        prs, "Advisory, or ordinary bug?",
        ("security", "Gets an advisory and a CVE",
         "Anything that crosses a boundary: pre-auth reachable, reads memory the "
         "client should not see, or corrupts memory with a path to code execution."),
        ("reliability", "Fixed in the normal release process",
         "Availability-only bugs that need an authenticated attacker. Still real "
         "bugs, still fixed, no embargo and no coordination."),
        (CARD_A, CARD_B),
    ), "Historically we issued a CVE for anything touching availability, "
       "confidentiality or integrity. That swept in an authenticated client crashing "
       "the server with a malformed command, and we coordinated patches for it.")

    # --- 10. the reasoning behind the line --------------------------------------
    notes(slide_from_layout(prs, LAYOUT_INSIGHT, [
        "A client that can run EVAL can already pin a core forever.",
        "So a bug that hangs the server hands it nothing it did not have. No "
        "advisory, no embargo, no coordination.",
    ]), "This is the slide people argue with, so give it room. The test is not 'is "
        "this bad' but 'does this give the attacker something new'. Someone will ask "
        "about shared hosting and untrusted tenants: the answer is that if untrusted "
        "clients hold EVAL, the deployment is already the problem.")

    # --- 11. the audit pipeline -------------------------------------------------
    notes(slide_three_box(
        prs, "How the adversarial audits work",
        [("observability", "One model proposes",
          "It reads a subsystem, a threat model, or a single feature, and writes "
          "down candidate bugs."),
         ("hardening", "A second model argues back",
          "Reading the same source, it tries to disprove each candidate, then "
          "verifies and reproduces whatever survives."),
         ("community", "A person sees the survivors",
          "A candidate reaches a maintainer only once it arrives with a test that "
          "produces a reproducible crash.")],
        NEUTRAL,
    ), "Harkrishn takes over here. The adversarial second stage is the part that "
       "makes this usable: most candidates die there, either because the code "
       "already handles the case the first model misread or because the impact is "
       "too small to act on. Without it you have built a slop generator aimed at "
       "yourself.")

    # --- 12. what it found ------------------------------------------------------
    notes(slide_from_layout(prs, LAYOUT_BIG_NUMBER, [
        "34",
        "real bugs found before anyone reported them",
        "Across Valkey and the JSON, search and bloom modules. " + BLOG,
    ]), "Run periodically across several frontier models, because they do not find "
        "the same things. The goal is to run this before every release so these "
        "never reach production, which is the only version of this that actually "
        "reduces the report volume.")

    # --- 13. the worked example -------------------------------------------------
    notes(slide_from_layout(prs, LAYOUT_LIST_ROWS, [
        "A TLS bug a human would miss",
        "TLS hands Valkey more bytes than one command, so a connection holds unread data",
        "Valkey walks the list of those connections once per event loop pass",
        "Walking a list means holding a pointer to the next connection",
        "The current connection's command is CLIENT KILL, aimed at that next connection",
        "The next iteration follows a pointer into freed memory",
    ]), "Walk it one row at a time. Every individual step is reasonable code, which "
        "is exactly why review missed it: the bug only exists in the interaction "
        "between the TLS buffering and the kill path. This affected every version of "
        "Valkey.")

    # --- 14. crash versus CVE ---------------------------------------------------
    notes(slide_from_layout(prs, LAYOUT_INSIGHT, [
        "The freed memory held an object the server calls through.",
        "An attacker who can place their own bytes there has a credible path to code "
        "execution. That is the line between a crash and a CVE.",
    ]), "Ties straight back to slide 9. We did not issue this CVE because the server "
        "crashed. We issued it because of what could be written into the hole.")

    # --- 15. shipping it --------------------------------------------------------
    notes(slide_two_box(
        prs, "The same fix, a year apart",
        ("migration", "A year ago",
         "A maintainer cherry-picks the fix into every supported branch by hand, "
         "resolving each conflict, while the embargo clock runs."),
        ("performance", "Now",
         "Agents open the backport pull requests and resolve the conflicts. The "
         "triage and release steps are codified as prompts."),
        (CARD_A, CARD_B),
    ), "The TLS bug is the example: it affected every supported version. Codifying "
       "the process as prompts matters as much as the automation, because whoever is "
       "shepherding a release now has the steps in front of them instead of in "
       "someone else's head.")

    # --- 16. what the automation did and did not change -------------------------
    # These were two consecutive statement slides. As one two-up they argue against
    # each other, which is the actual point, and it costs a slide less.
    notes(slide_two_box(
        prs, "What the automation bought, and what it did not",
        ("cost-efficiency", "What it bought",
         "Wake up to a zero-day, ship the fix the same day. That is the whole return, "
         "and it is the one maintainers actually want."),
        ("community", "What it did not change",
         "A human is still responsible for every merge. Keeping people in the loop "
         "matters more in a security release, not less."),
        (CARD_A, CARD_B),
    ), "Do not skip the right-hand card to save time. Half the room's objection to "
       "everything in the last twenty minutes is answered there.")

    # --- 17. what the audience does next ---------------------------------------
    notes(slide_three_box(
        prs, "What to do on Monday",
        [("reliability", "If you run it in production",
          "Move to the current patch release of every project you run, and build "
          "something that consumes new versions regularly."),
         ("security", "If you find a bug",
          "Mail security@lists.valkey.io rather than opening an issue, and include a "
          "reproducer that runs against the build you targeted."),
         ("hardening", "If you maintain a project",
          "Draw the line at what gives an attacker something new, then point the "
          "same tools at your own code first.")],
        NEUTRAL,
    ), "Three audiences, one action each. For the production row, also point at the "
       "deployment hardening guide: valkey.io/blog/properly-secure-your-valkey-"
       "deployment/. For the bug row, the reproducer is the ask that saves us the "
       "most time.")

    # --- 18. close --------------------------------------------------------------
    notes(slide_from_layout(prs, LAYOUT_INSIGHT, [
        "Bugs got cheaper to find, and they keep getting cheaper.",
        "Putting the same tools in maintainers' hands is what keeps projects secure "
        "for the long term.",
    ]), "The honest close: we are not claiming to have solved this. The wave is not "
        "going away, and the answer is not asking researchers to stop. It is closing "
        "the tooling gap between the people finding bugs and the people fixing them.")

    # --- 19. thank you ----------------------------------------------------------
    notes(slide_from_layout(prs, LAYOUT_THANKS, [
        "Thank you",
        "Madelyn Olson  ·  Harkrishn Patro",
        "security@lists.valkey.io",
        "valkey.io/blog/keeping-up-with-ai-valkey-security/",
    ]), "Leave this up for questions. Expect: how do you stop researchers gaming the "
        "new policy, what do the audit prompts look like, and who pays for the model "
        "calls.")

    prs.save(OUT + ".building")
    Presentation(OUT + ".building")     # reopen: if it raises, PowerPoint would too
    os.replace(OUT + ".building", OUT)
    print(f"wrote {OUT}")

    check = Presentation(OUT)
    print(f"verify: {len(check.slides)} slides, {len(check.slide_layouts)} layouts")
    for i, s in enumerate(check.slides, 1):
        head = next((sh.text_frame.text.split("\n")[0][:52] for sh in s.shapes
                     if sh.has_text_frame and sh.text_frame.text.strip()), "")
        has_notes = bool(s.notes_slide.notes_text_frame.text.strip())
        print(f"  {i:2} {s.slide_layout.name:28} notes={'y' if has_notes else 'NO'}  {head}")
    return OUT


if __name__ == "__main__":
    build()
