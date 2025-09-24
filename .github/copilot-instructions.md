# GitHub Copilot Instructions for win32-taskscheduler

This document provides comprehensive instructions for GitHub Copilot when working with the win32-taskscheduler repository.

## Repository Overview

The `win32-taskscheduler` is a Ruby library that provides an interface to the MS Windows Task Scheduler. It is analogous to the Unix cron daemon and allows creation, configuration, and deletion of scheduled tasks on Windows systems.

### Repository Structure

```
win32-taskscheduler/
├── .github/                      # GitHub configuration and templates
│   ├── workflows/                # GitHub Actions workflows
│   │   └── lint.yml             # Linting and spellcheck workflow
│   ├── CODEOWNERS               # Code ownership definitions
│   ├── ISSUE_TEMPLATE.md        # Issue template
│   ├── PULL_REQUEST_TEMPLATE.md # PR template
│   └── dependabot.yml           # Dependabot configuration
├── lib/                         # Main library code
│   ├── win32-taskscheduler.rb   # Main entry point
│   └── win32/                   # Core implementation
│       ├── taskscheduler.rb     # Main TaskScheduler class
│       └── taskscheduler/       # Supporting modules
│           ├── constants.rb     # Constants definitions
│           ├── helper.rb        # Helper utilities
│           ├── sid.rb           # Security identifier handling
│           ├── time_calc_helper.rb # Time calculation utilities
│           └── version.rb       # Version information
├── spec/                        # RSpec test suite
│   ├── spec_helper.rb           # RSpec configuration
│   ├── functional/              # Functional tests
│   └── unit/                    # Unit tests
├── test/                        # Test-unit test suite
│   └── test_taskscheduler.rb    # Legacy test suite
├── examples/                    # Usage examples
│   └── taskscheduler_example.rb # Example implementation
├── vendor/                      # Bundled dependencies
├── Gemfile                      # Ruby dependencies
├── Rakefile                     # Build and task definitions
├── README.md                    # Project documentation
├── CHANGELOG.md                 # Version history
├── CODE_OF_CONDUCT.md          # Code of conduct
├── LICENSE                      # Artistic 2.0 license
└── win32-taskscheduler.gemspec  # Gem specification
```

## Key Technologies and Dependencies

- **Language**: Ruby (3.1+)
- **Main Dependencies**: 
  - `ffi` - Foreign Function Interface for Windows API calls
  - `structured_warnings` - Enhanced warning system
- **Development Dependencies**:
  - `test-unit` - Unit testing framework
  - `win32-security` - Windows security utilities
  - `chefstyle` - Ruby code style checker
  - `rspec` - Behavior-driven development testing framework

## Jira Integration Workflow

When a Jira ID is provided in task requests:

1. **Use the atlassian-mcp-server MCP server** to fetch Jira issue details
2. **Read and analyze** the story/task description thoroughly
3. **Implement the required functionality** based on the Jira requirements
4. **Follow the complete development workflow** outlined below

### MCP Server Commands for Jira Integration

```bash
# Get accessible Atlassian resources
mcp_atlassian-mcp_getAccessibleAtlassianResources

# Search for issues
mcp_atlassian-mcp_search --query "your search terms"

# Get specific issue details
mcp_atlassian-mcp_getJiraIssue --cloudId <cloudId> --issueIdOrKey <issue-key>

# Add comments to issues
mcp_atlassian-mcp_addCommentToJiraIssue --cloudId <cloudId> --issueIdOrKey <issue-key> --commentBody "your comment"
```

## Testing Requirements

### Coverage Standards
- **Minimum test coverage**: >80% for all new code
- **Test both unit and functional scenarios**
- **Include edge cases and error conditions**

### Testing Frameworks
- **Primary**: RSpec (`spec/` directory)
- **Legacy**: test-unit (`test/` directory)

### Running Tests
```bash
# Run RSpec tests
bundle exec rake spec

# Run all tests (includes linting)
bundle exec rake

# Run linting only
bundle exec rake style

# Check test coverage
bundle exec rake coverage  # If coverage tools are configured
```

### Test Creation Guidelines
1. Create unit tests in `spec/unit/` mirroring the `lib/` structure
2. Create functional tests in `spec/functional/`
3. Follow existing test patterns and naming conventions
4. Test Windows-specific functionality thoroughly
5. Mock Windows API calls where appropriate for cross-platform development

## Pull Request Workflow

### Branch Creation and Management
```bash
# Create and switch to new branch using Jira ID as branch name
gh auth status  # Check authentication status
git checkout -b <JIRA-ID>

# Make your changes and commit with DCO sign-off
git add .
git commit -s -m "feat: implement <feature> for <JIRA-ID>"

# Push branch to origin
git push -u origin <JIRA-ID>
```

### Creating Pull Requests
```bash
# Create PR using GitHub CLI
gh pr create \
  --title "feat: <Brief description> - <JIRA-ID>" \
  --body "$(cat << 'EOF'
<h2>Description</h2>
<p>Brief summary of changes made</p>

<h2>Issues Resolved</h2>
<p>Resolves: <JIRA-ID></p>

<h2>Changes Made</h2>
<ul>
<li>Change 1</li>
<li>Change 2</li>
<li>Change 3</li>
</ul>

<h2>Testing</h2>
<p>All tests pass with >80% coverage</p>

<h2>DCO Compliance</h2>
<p>All commits have been signed-off for the Developer Certificate of Origin</p>
EOF
)" \
  --label "Aspect: Documentation,Aspect: Testing" \
  --assignee @me
```

## DCO Compliance Requirements

### Developer Certificate of Origin (DCO)
All commits **MUST** be signed-off to certify compliance with the Developer Certificate of Origin.

**Required for every commit**:
```bash
git commit -s -m "your commit message"
```

The `-s` flag adds the required `Signed-off-by` line with your name and email.

### DCO Sign-off Format
```
Signed-off-by: Your Name <your.email@example.com>
```

### DCO Verification
- All PRs are automatically checked for DCO compliance
- Missing DCO sign-offs will block PR merging
- Use `git commit --amend -s` to add DCO to the last commit if forgotten

## GitHub Workflows and CI/CD

### Available GitHub Actions
The repository uses GitHub Actions for continuous integration:

#### Lint Workflow (`.github/workflows/lint.yml`)
- **Triggers**: Pull requests and pushes to main branch
- **Jobs**:
  - **ChefStyle**: Ruby code style checking using `chefstyle`
  - **Spellcheck**: Documentation spellcheck using `cspell`
- **Ruby Version**: 3.1
- **Platform**: Ubuntu Latest

### Build System Integration

#### Expeditor Labels
The repository uses Expeditor for automated release management:

- `Expeditor: Bump Version Major` - Triggers major version bump
- `Expeditor: Bump Version Minor` - Triggers minor version bump  
- `Expeditor: Skip All` - Skips all merge actions
- `Expeditor: Skip Changelog` - Skips changelog update
- `Expeditor: Skip Version Bump` - Skips version bumping

### Available Repository Labels

#### Aspect Labels
- `Aspect: Documentation` - Documentation improvements
- `Aspect: Integration` - Integration with other systems
- `Aspect: Packaging` - Distribution and packaging
- `Aspect: Performance` - Performance improvements
- `Aspect: Portability` - Cross-platform compatibility
- `Aspect: Security` - Security-related changes
- `Aspect: Stability` - Stability improvements
- `Aspect: Testing` - Test coverage and CI
- `Aspect: UI` - User interface changes
- `Aspect: UX` - User experience improvements

#### Platform Labels
- `Platform: AWS`, `Platform: Azure`, `Platform: GCP` - Cloud platforms
- `Platform: Docker` - Container support
- `Platform: Linux`, `Platform: macOS` - Operating systems
- `Platform: Debian-like`, `Platform: RHEL-like`, `Platform: SLES-like` - Linux distributions

#### Special Labels
- `dependencies` - Dependency updates
- `hacktoberfest-accepted` - Hacktoberfest contributions
- `oss-standards` - OSS standardization

## Complete Development Workflow

### Step-by-Step Process

#### 1. Initial Setup and Planning
1. **Receive task with Jira ID**
2. **Fetch Jira issue details** using MCP server
3. **Analyze requirements** and break down into implementation steps
4. **Plan testing approach** ensuring >80% coverage target

#### 2. Branch and Development Setup
```bash
# Authenticate with GitHub (if needed)
gh auth status

# Create feature branch using Jira ID
git checkout -b <JIRA-ID>

# Ensure dependencies are installed
bundle install
```

#### 3. Implementation Phase
1. **Implement core functionality** in `lib/` directory
2. **Follow existing code patterns** and Ruby conventions
3. **Handle Windows-specific requirements** appropriately
4. **Add appropriate error handling** and logging

#### 4. Testing Phase
1. **Create comprehensive unit tests** in `spec/unit/`
2. **Create functional tests** in `spec/functional/` if needed
3. **Ensure test coverage >80%**
4. **Run all tests locally**:
   ```bash
   bundle exec rake spec
   bundle exec rake style
   ```

#### 5. Documentation and Examples
1. **Update code documentation** with YARD format
2. **Add usage examples** in `examples/` if appropriate  
3. **Update README.md** if public API changes
4. **Update CHANGELOG.md** following existing format

#### 6. Pre-commit Verification
1. **Run full test suite**: `bundle exec rake`
2. **Verify DCO compliance** on all commits
3. **Check code style** passes ChefStyle requirements
4. **Ensure no prohibited files** are modified

#### 7. Commit and Push
```bash
# Stage changes
git add .

# Commit with DCO sign-off
git commit -s -m "feat: implement <feature-description> for <JIRA-ID>"

# Push to origin
git push -u origin <JIRA-ID>
```

#### 8. Pull Request Creation
```bash
# Create PR with comprehensive description
gh pr create \
  --title "feat: <Brief description> - <JIRA-ID>" \
  --body "<HTML formatted description with changes summary>" \
  --label "Aspect: Testing,<other-relevant-labels>"
```

#### 9. Post-PR Activities
1. **Monitor CI/CD pipeline** results
2. **Address any review feedback**
3. **Update Jira issue** with PR link and status
4. **Respond to automated checks** and fix any issues

## Prompt-Based Development Guidelines

### After Each Step Provide:
1. **Clear summary** of what was accomplished
2. **Current status** of the overall task
3. **Next step** in the workflow
4. **Remaining steps** to complete the task
5. **Ask for confirmation** to continue with the next step

### Example Progress Update Format:
```
✅ **Completed**: [Step description]
📋 **Current Status**: [Overall progress summary]
🔄 **Next Step**: [Specific next action]
📝 **Remaining Steps**: 
   - Step 1
   - Step 2
   - Step 3

❓ **Continue with [next step]?** (y/n)
```

### Confirmation Pattern
Always ask before proceeding to the next major step:
- "Should I proceed with implementing the core functionality?"
- "Ready to create the unit tests for this feature?"
- "Shall I create the pull request now?"

## File Modification Restrictions

### Protected Files (Do Not Modify)
- `.github/workflows/` - CI/CD configurations
- `Gemfile.lock` - Dependency lock file (unless adding dependencies)
- `vendor/` - Bundled dependencies
- `.gitignore` - Version control ignore rules
- License files

### Require Special Attention
- `win32-taskscheduler.gemspec` - Only modify for dependency changes
- `VERSION` - Only modify for version bumps
- `Rakefile` - Only modify for build process improvements
- `Gemfile` - Only modify when adding new dependencies

## Code Style and Standards

### Ruby Style Guidelines
- Follow **ChefStyle** conventions (based on RuboCop)
- Use **2-space indentation**
- Follow **semantic versioning** for changes
- Write **descriptive commit messages**
- Use **YARD documentation format** for code documentation

### Naming Conventions
- Use `snake_case` for methods and variables
- Use `SCREAMING_SNAKE_CASE` for constants
- Use descriptive names that reflect Windows Task Scheduler terminology
- Follow existing patterns in the codebase

### Error Handling
- Provide clear error messages
- Handle Windows-specific exceptions appropriately
- Log important operations and failures
- Use structured warnings where appropriate

## Security Considerations

### Windows Security Context
- Handle Windows security identifiers (SIDs) properly
- Respect Windows user permissions and privileges
- Validate task scheduler access rights
- Sanitize user inputs for task creation

### Dependency Management
- Keep dependencies up to date
- Review security implications of new dependencies
- Use Dependabot for automated dependency updates
- Follow Chef's security guidelines

This comprehensive guide ensures consistent, high-quality development practices while maintaining the repository's standards and Chef's development workflow requirements.
