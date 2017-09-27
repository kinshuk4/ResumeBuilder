import yaml

def yaml2dict(filename):
    with open(filename, "r") as stream:
        resume_dict = yaml.load(stream)

    return resume_dict

def main():
    resumeFile = "../demo/sample-resume.yaml"
    resume_dict = yaml2dict(resumeFile)
    print(resume_dict)

if __name__ == '__main__':
    main()

