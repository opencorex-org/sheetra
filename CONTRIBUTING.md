# Contributing to Sheetra

Thank you for considering contributing to Sheetra! We welcome contributions from the community to help improve and grow the project. This document outlines the guidelines for contributing to ensure a smooth and collaborative process.

---

## Code of Conduct
By participating in this project, you agree to abide by our [Code of Conduct](CODE_OF_CONDUCT.md). Please ensure that your interactions are respectful and inclusive.

---

## How Can I Contribute?

### Reporting Bugs
If you find a bug, please create an issue in the [GitHub repository](https://github.com/opencorex-org/sheetra/issues) with the following details:
- A clear and descriptive title.
- Steps to reproduce the issue.
- Expected and actual behavior.
- Screenshots or logs, if applicable.

### Suggesting Features
We welcome feature requests! To suggest a feature, open an issue in the [GitHub repository](https://github.com/opencorex-org/sheetra/issues) and include:
- A detailed description of the feature.
- The problem it solves or the use case.
- Any examples or references.

### Submitting Code Changes
1. Fork the repository and create a new branch for your changes.
2. Ensure your branch name is descriptive (e.g., `fix-bug-123` or `add-new-feature`).
3. Make your changes and write clear, concise commit messages.
4. Test your changes thoroughly.
5. Submit a pull request (PR) to the `main` branch with a detailed description of your changes.

---

## Development Workflow

### Setting Up the Project
1. Clone the repository:
   ```bash
   git clone https://github.com/opencorex-org/sheetra.git
   ```
2. Install dependencies using `pnpm`:
   ```bash
   pnpm install
   ```
3. Run the development server:
   ```bash
   pnpm dev
   ```

### Running Tests
Run the test suite to ensure your changes do not break existing functionality:
```bash
pnpm test
```

### Building the Project
To build the project for production:
```bash
pnpm build
```

---

## Style Guidelines

### Code Style
- Follow the existing code style and structure.
- Use TypeScript for all new code.
- Run the linter before submitting your changes:
  ```bash
  pnpm lint
  ```

### Documentation
- Update relevant documentation in the `docs/` folder for any changes.
- Ensure your code is well-commented.

---

## Commit Message Guidelines
We follow the [Conventional Commits](https://www.conventionalcommits.org/) specification. Examples:
- `feat: add support for new feature`
- `fix: resolve issue with workbook creation`
- `docs: update API documentation`

---

## License
By contributing, you agree that your contributions will be licensed under the [MIT License](../LICENSE).

---

Thank you for contributing to Sheetra! Together, we can make it even better.