import datetime
import sys
import pathlib

from dictionary_generator.table.generator import Generator


def main():
    name = sys.argv[1]
    date_of_birth = sys.argv[2]
    group = sys.argv[3]
    weight = sys.argv[4]
    height = sys.argv[5]

    document_name = sys.argv[9]
    start_str = sys.argv[6];
    start = datetime.datetime.strptime(start_str, "%d.%m.%Y")

    end_str = sys.argv[7];
    end = datetime.datetime.strptime(end_str, "%d.%m.%Y")


    frequency = int(sys.argv[8])

    generator = Generator()

    document = generator.generate(
        name=name,
        date_of_birth=date_of_birth,
        group=group,
        weight=weight,
        height=height,
        start=start,
        end=end,
        frequency=frequency,
    )

    here = pathlib.Path(__file__).parent
    document_path = here / (document_name + ".docx")

    with open(document_path, "wb") as w:
        w.write(document.getbuffer().tobytes())

    print("Документ успешно сохранен в ", document_path)


if __name__ == "__main__":
    main()
