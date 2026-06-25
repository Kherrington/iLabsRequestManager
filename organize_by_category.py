#!/usr/bin/env python3
"""
Organize markdown files into category folders based on content analysis.
"""

import shutil
from pathlib import Path

# Category mapping based on wiki structure
CATEGORIES = {
    "microscopes": [
        "517178903.md",  # TIRF/N-STORM
        "517182052.md",  # N-SIM
        "517182054.md",  # CSU-W1
        "517182081.md",  # CAMM #1
        "517182083.md",  # CAMM #2
        "522589638.md",  # Crest LFOV
        "6D_517182048.md",  # 6D
        "711464117.md",  # Snouty
        "AZ100-Light-Sheet_517182060.md",
        "C-Trap-Optical-Tweezers_698319820.md",
        "CVRI-Leica-SPE-Confocal_517182073.md",
        "CVRI-Nikon-Epifluorescence-Microscope_517182069.md",
        "CVRI-Nikon-Spinning-Disk-Confocal_517182067.md",
        "OMX-SR_517182062.md",
        "QLIPP-Timelapse-with-Nanoindentor_621616516.md",
        "Spinning-Disk-Confocal_517178889.md",
        "Time-Lapse_517182042.md",
        "TruLive3D-Light-Sheet_674609376.md",
        "Upright-Spinning-Disk-Instructions_654283421.md",
        "Weill---Leica-Widefield_561953534.md",
        "Weill--CSU-W1-SoRA-Spinning-Disk-Confocal_561953513.md",
        "Weill--Leica-Fluorescence-Stereo_561953591.md",
        "Weill--Leica-Thunder-Stereo_561953587.md",
        "Weill--Molecular-Devices-Image-Express-Confocal-HTai_561953668.md",
        "Weill--Stellaris-8-Tau-STED_561953181.md",
        "Other-Equipment-and-Objectives_517182065.md",
        "Microscopes_517178884.md",  # Main category page
    ],
    "data-analysis": [
        "Data-Analysis_517182106.md",  # Main category page
        "Data-Analysis-Workstations_517178905.md",
        "NIC-Analysis-workstations_685901416.md",
        "Data-Storage-and-Compute_685901383.md",
        "The-previous-NIC-Data-server_685901387.md",
        "720372404.md",  # Remote Access
        "Add-a-scale-bar-in-Fiji_634298063.md",
        "Fiji--Viewing-orthagonal-slices_701468027.md",
        "Finding-your-pixel-size_634298046.md",
        "How-to-Acquire-Flat-Field-correction_517182172.md",
        "Stitching-Images-Acquired-in-Micro-Manager_517182166.md",
        "C-Trap-Data-Analysis_698319889.md",
        "Cilia-Analysis_689317093.md",
        "ClearVolume_701467993.md",
        "Huygens_517184442.md",
        "ImageJ-and-Variants_517184301.md",
        "MATLAB_517184427.md",
        "Python-Scripts-for-TruLive_698323118.md",
    ],
    "sample-preparation": [
        "Sample-Preparation_517182102.md",  # Main category page
        "689319159.md",  # ATP/ADP Reporter
        "Brain-Slice-Preparation_647401317.md",
        "Calcium-FURA2-imaging_517184507.md",
        "Clearing-methods_517184494.md",
        "Cy7-and-750nm-Range-Dyes_621613530.md",
        "Cytoskeleton-Fixation-Referances_654278900.md",
        "Dyes-for-405nm-excitation_517184492.md",
        "Fluorescent-Proteins-for-Localization-Microscopy_517197198.md",
        "High-Precision-Coverslips_517184498.md",
        "Make-your-own-flow-chamber_689313954.md",
        "Mitochondrial-Dyes_574138207.md",
        "Plates-and-dishes-for-imaging_517184496.md",
        "Protocols-for-Immunocytochemistry-in-cell-culture_517184505.md",
        "Sample-preparation-for-SIM-imaging_517184501.md",
        "STORM-sample-preparation-and-imaging_517184503.md",
    ],
    "references-and-education": [
        "Microscopy-References-and-Education_517178908.md",  # Main category page
        "517182143.md",  # Books, Websites
        "517193525.md",  # 2012 Course
        "517197331.md",  # Spring 2013 Course
        "517198168.md",  # 2013 Course
        "517198186.md",  # 2014 Course
        "Microscopy-Courses_517182136.md",
        "Monthly-Microscopy-Users-Meeting_520586430.md",
        "Presentations_517182141.md",
        "Examples-of-Microscope-Methods_641107539.md",
        "Methods-Examples_666378595.md",
    ],
    "calm-information": [
        "Other-CALM-Information_517182095.md",  # Main category page
        "Acknowledgements_599783507.md",
        "iLab_701465591.md",
        "User-Quick-Start-Guides_517182146.md",
    ],
    "miscellaneous": [
        "Miscellaneous_517182118.md",  # Main category page
        "Laser-Cut-and-3D-Printed-Parts-for-Microscopy_517182120.md",
        "Outside-fabrication-resources_517182124.md",
        "PSFs-and-aberrations_517182122.md",
    ],
}

def organize_files():
    """Organize markdown files into category folders.

    Creates category subdirectories and copies files from the main pages
    directory into their appropriate category folders.
    """
    pages_dir = Path("wiki/pages")

    # Validate pages directory exists
    if not pages_dir.exists():
        print(f"[ERROR] Pages directory not found: {pages_dir}")
        return

    # Track which files have been categorized
    categorized = set()
    missing_files = []

    for category, files in CATEGORIES.items():
        # Create category directory
        category_dir = pages_dir / category
        category_dir.mkdir(exist_ok=True)

        print(f"\n[{category.upper()}]")

        for filename in files:
            source = pages_dir / filename
            dest = category_dir / filename

            if source.exists():
                try:
                    # Copy (not move) so we can verify
                    shutil.copy2(source, dest)
                    categorized.add(filename)
                    print(f"  + {filename}")
                except Exception as e:
                    print(f"  ! Error copying {filename}: {e}")
            else:
                print(f"  ! Missing: {filename}")
                missing_files.append((category, filename))

    # Check for uncategorized files
    print("\n[UNCATEGORIZED FILES]")
    all_files = set(f.name for f in pages_dir.glob("*.md"))
    uncategorized = all_files - categorized - {"_file_mapping.txt"}

    if uncategorized:
        for filename in sorted(uncategorized):
            print(f"  ? {filename}")
    else:
        print("  (none)")

    print(f"\n[SUMMARY]")
    print(f"  Total files: {len(all_files)}")
    print(f"  Categorized: {len(categorized)}")
    print(f"  Uncategorized: {len(uncategorized)}")
    if missing_files:
        print(f"  Missing: {len(missing_files)}")

    return {
        'total': len(all_files),
        'categorized': len(categorized),
        'uncategorized': len(uncategorized),
        'missing': len(missing_files)
    }

if __name__ == "__main__":
    organize_files()
