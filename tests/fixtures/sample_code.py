"""Sample working python-pptx code for testing."""

from pptx import Presentation
from pathlib import Path


def create_test_presentation():
    """Create a simple test presentation."""
    prs = Presentation()

    # Slide 1: Title
    slide1 = prs.slides.add_slide(prs.slide_layouts[0])
    slide1.shapes.title.text = "Test Presentation"
    slide1.placeholders[1].text = "A simple test"

    # Slide 2: Content
    slide2 = prs.slides.add_slide(prs.slide_layouts[1])
    slide2.shapes.title.text = "Key Points"
    body = slide2.placeholders[1]
    tf = body.text_frame
    tf.text = "Point 1"
    tf.add_paragraph().text = "Point 2"
    tf.add_paragraph().text = "Point 3"

    # Slide 3: Conclusion
    slide3 = prs.slides.add_slide(prs.slide_layouts[1])
    slide3.shapes.title.text = "Thank You"
    slide3.placeholders[1].text = "Final remarks and contact information."

    # Save
    output_path = Path('.output') / 'test_fixture.pptx'
    output_path.parent.mkdir(exist_ok=True)
    prs.save(str(output_path))
    print(f"Presentation saved to {output_path}")


if __name__ == '__main__':
    create_test_presentation()
