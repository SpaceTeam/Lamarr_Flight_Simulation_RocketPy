# TODO

Open tasks for the RocketPy flight simulation. 


## General

## Code quality

- Add type hints across the codebase.
- Add unit tests.
- Unify docstrings and mirror RocketPy's convention.

### Simulation features

- Result sanity checks: warn when the stability margin is too low or the simulation produces other implausible values. Log a possible reason.
- Standalone vs. variation check: compare a standalone simulation with the variation simulation using single values to make sure both return the same results. Standalone simulations are still needed for sending them to launch day organisators.
- Standardize one solid and one liquid standalone version. 
- Add variation for fin shape points.

### Streamlit app

- Per-project RocketPy version: switching projects in Streamlit does not switch to the RocketPy version of the project.
  Idea: store the required RocketPy version in each project's `config.toml` and install it when the user clicks Run.

### OpenRocket<>RocketPy

- Write a script that converts an OpenRocket config (XML file) to a RocketPy config and vice versa. It prompts the user for fields that the other tool does not have or handles differently. It also can compare existing config files and flag differences.
- Write a script that compares simulation results of both tools.


## Project-specific

### ALBATROSS

_No open tasks._

### ALBATROSS II

- Update config values marked with TODO.

### CANSAT

_No open tasks._

### EUROC2025

_No open tasks._

### FIRST

_No open tasks._
