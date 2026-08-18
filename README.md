# College Music Syllabi and Lesson Plans

Editable teaching packages for college-level music classes in the Philippines.

## Deliverables

- `Music_Curriculum_Master_Syllabus_PH.docx` - combined master syllabus
- `Music_Fundamentals_Syllabus_and_Lesson_Plans.docx`
- `Music_Theory_I_Syllabus_and_Lesson_Plans.docx`
- `Music_Theory_II_Syllabus_and_Lesson_Plans.docx`
- `Sight_Singing_Syllabus_and_Lesson_Plans.docx`

Each separate course package includes an 18-week syllabus, detailed weekly lesson plans, activities, assessment rubrics, policies, differentiation guidance, and visual aids.

## Source materials used by the instructor

- Alfred's *Essentials of Music Theory* by Andrew Surmani, Karen Farnum Surmani, and Morton Manus
- Benner, *Theory for Piano Students*, G. Schirmer
- Audrey Snyder, *The Sight Singer*, Volume 1

Exact textbook pages remain instructor-assigned because pagination can differ by edition or printing.

## Rebuilding

The Python builders use `python-docx` and Pillow. Run `build_syllabus.py` first to generate the master file and visual assets, then run `build_separate_course_packages.py` to generate the four course packages.
