"""Slide operations: clone, move, transitions, backgrounds, and notes.

python-pptx cannot clone slides, move slides, set transitions, or
manipulate slide backgrounds. offidized handles all of these natively.
"""

from offidized import Presentation, SlideTransition

EMU = 914400

pres = Presentation()
pres.set_slide_size(12192000, 6858000)


def add_titled_slide(title: str, color: str) -> int:
    idx = pres.add_slide()
    slide = pres.get_slide(idx)
    slide.set_background_solid(color)
    s = slide.add_shape(title)
    shape = slide.get_shape(s)
    shape.set_geometry(EMU, int(2.5 * EMU), 10 * EMU, 2 * EMU)
    para = shape.add_paragraph_with_text(title)
    run = para.get_run(0)
    run.set_font_color("FFFFFF")
    run.set_font_size(4400)
    run.set_bold(True)
    para.set_alignment("ctr")
    return idx


# Create slides with different backgrounds
idx1 = add_titled_slide("Introduction", "2F5597")
idx2 = add_titled_slide("Main Content", "548235")
idx3 = add_titled_slide("Conclusion", "C0504D")

# Set transitions (not possible with python-pptx)
pres.get_slide(idx1).set_transition(SlideTransition("fade"))
pres.get_slide(idx2).set_transition(SlideTransition("push"))
pres.get_slide(idx3).set_transition(SlideTransition("wipe"))

# Add speaker notes (limited in python-pptx)
pres.get_slide(idx1).set_notes("Welcome the audience and introduce the topic.")
pres.get_slide(idx2).set_notes("Cover the key points. Pause for questions.")
pres.get_slide(idx3).set_notes("Summarize and provide call to action.")

# Clone slide 2 (not possible with python-pptx)
cloned_idx = pres.clone_slide(idx2)
cloned = pres.get_slide(cloned_idx)
cloned.set_background_solid("7030A0")
# Update the cloned slide's text
shape = cloned.get_shape(0)
para = shape.get_paragraph(0)
run = para.get_run(0)
run.set_text("Bonus Slide (Cloned!)")

# Move the cloned slide to position 2 (not possible with python-pptx)
pres.move_slide(cloned_idx, 2)

# Presentation-wide find/replace (not possible with python-pptx)
count = pres.replace_text("Conclusion", "Final Thoughts")
print(f"Replaced {count} occurrence(s) of 'Conclusion'")

print(f"Total slides: {pres.slide_count()}")
pres.save("02_slide_operations.pptx")
print("Created 02_slide_operations.pptx")
